// This file contains functions responsible for archiving old clinic data
// and associated participant responses.

function runManualArchive() {
  archiveOldClinics(true);
}

function runDailyArchive() {
  archiveOldClinics(false);
}

function archiveOldClinics(isManualTrigger) {
  const logPrefix = isManualTrigger ? "Handmatige archivering" : "Automatische dagelijkse archivering";
  logMessage(`----- START ${logPrefix} -----`);

  try {
    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const dataClinicsSheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);

    // Create archive sheet if it doesn't exist
    let archiveSheet = dataClinicsSpreadsheet.getSheetByName(ARCHIVE_SHEET_NAME);
    if (!archiveSheet) {
      archiveSheet = dataClinicsSpreadsheet.insertSheet(ARCHIVE_SHEET_NAME);
      const headers = dataClinicsSheet.getRange(1, 1, 1, dataClinicsSheet.getLastColumn()).getValues();
      archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
      logMessage(`Archief sheet '${ARCHIVE_SHEET_NAME}' aangemaakt.`);
    }

    // Define the threshold for archiving (e.g., 30 days old)
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    thirtyDaysAgo.setHours(0, 0, 0, 0); // Normalize to start of day

    const allData = dataClinicsSheet.getDataRange().getValues();
    if (allData.length <= DATA_CLINICS_START_ROW -1) { // Only headers or empty
        logMessage(`Geen clinics gevonden om te archiveren in '${DATA_CLINICS_SHEET_NAME}'.`);
        if (isManualTrigger) SpreadsheetApp.getUi().alert('Geen clinics gevonden om te archiveren.');
        return;
    }
    const headers = allData.shift(); // Remove headers for data processing

    const clinicsToKeep = [];
    const clinicsToArchive = [];
    const archivedClinicNames = new Set(); // Store names of clinics that are archived

    allData.forEach(row => {
      const dateValue = row[DATE_COLUMN_INDEX - 1];
      if (!dateValue) {
        clinicsToKeep.push(row); // Keep rows without a date
        return;
      }
      const clinicDate = new Date(dateValue);
      if (isNaN(clinicDate.getTime())) {
        clinicsToKeep.push(row); // Keep rows with invalid dates
        logMessage(`WAARSCHUWING: Ongeldige datum gevonden op rij, overgeslagen voor archivering.`);
        return;
      }

      if (clinicDate < thirtyDaysAgo) {
        clinicsToArchive.push(row);
        const clinicName = `${getDutchDateString(clinicDate)} ${String(row[TIME_COLUMN_INDEX - 1] || '').trim()}, ${String(row[LOCATION_COLUMN_INDEX - 1] || '').trim()}`;
        archivedClinicNames.add(clinicName);
        logMessage(`Clinic "${clinicName}" gemarkeerd voor archivering.`);
      } else {
        clinicsToKeep.push(row);
      }
    });

    const numArchived = clinicsToArchive.length;
    if (numArchived === 0) {
      const message = 'Geen clinics ouder dan 30 dagen gevonden om te archiveren.';
      logMessage(message);
      if (isManualTrigger) SpreadsheetApp.getUi().alert(message);
      logMessage(`----- EINDE ${logPrefix} -----`);
      return;
    }

    // Append archived clinics to the archive sheet
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, numArchived, clinicsToArchive[0].length).setValues(clinicsToArchive);
    logMessage(`${numArchived} clinics verplaatst naar '${ARCHIVE_SHEET_NAME}'.`);

    // Clear and rewrite Data Clinics sheet
    dataClinicsSheet.getRange(DATA_CLINICS_START_ROW, 1, dataClinicsSheet.getLastRow(), dataClinicsSheet.getLastColumn()).clearContent();
    if (clinicsToKeep.length > 0) {
      dataClinicsSheet.getRange(DATA_CLINICS_START_ROW, 1, clinicsToKeep.length, clinicsToKeep[0].length).setValues(clinicsToKeep);
      logMessage(`${clinicsToKeep.length} clinics overgebleven in '${DATA_CLINICS_SHEET_NAME}'.`);
    } else {
        logMessage(`Alle clinics zijn gearchiveerd uit '${DATA_CLINICS_SHEET_NAME}'.`);
    }

    // Clean up participant response sheets
    const responseSs = SpreadsheetApp.getActiveSpreadsheet();
    [OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME].forEach(sheetName => {
      const responseSheet = responseSs.getSheetByName(sheetName);
      if (!responseSheet) {
        logMessage(`WAARSCHUWING: Respons-sheet '${sheetName}' niet gevonden voor opschonen.`);
        return;
      }
      
      const responseData = responseSheet.getDataRange().getValues();
      if (responseData.length < 2) return; // Only headers
      
      const responseHeaders = responseData.shift();
      const eventColIdx = responseHeaders.indexOf(FORM_EVENT_QUESTION_TITLE);
      
      if (eventColIdx === -1) {
        logMessage(`WAARSCHUWING: Kolom '${FORM_EVENT_QUESTION_TITLE}' ontbreekt in respons-sheet '${sheetName}'. Kan deze sheet niet opschonen.`);
        return;
      }
      
      const responsesToKeep = responseData.filter(row => {
          const participantClinicName = (row[eventColIdx] || '').replace(/\s\(.*\)$/, '').trim();
          return !archivedClinicNames.has(participantClinicName);
      });

      // Clear existing responses and write back the ones to keep
      responseSheet.getRange(2, 1, responseSheet.getLastRow(), responseSheet.getLastColumn()).clearContent();
      if (responsesToKeep.length > 0) {
        responseSheet.getRange(2, 1, responsesToKeep.length, responsesToKeep[0].length).setValues(responsesToKeep);
      }
      logMessage(`${responseData.length - responsesToKeep.length} reacties verwijderd uit '${sheetName}'.`);
    });

    const plural = numArchived === 1 ? 'clinic' : 'clinics';
    const successMessage = `${numArchived} ${plural} gearchiveerd.`;
    logMessage(successMessage);
    if (isManualTrigger) SpreadsheetApp.getUi().alert(successMessage);

  } catch (e) {
    const errorMessage = `FOUT tijdens archiveren: ${e.toString()}\n${e.stack}`;
    Logger.log(errorMessage);
    logMessage(errorMessage);
    if (isManualTrigger) SpreadsheetApp.getUi().alert(errorMessage);
  } finally {
    logMessage(`----- EINDE ${logPrefix} -----`);
    flushLogs(); // Ensures logs are always saved
  }
}