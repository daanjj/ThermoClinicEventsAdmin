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

    // Archive participant response data instead of deleting it
    const responseSs = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create participant archive sheet if it doesn't exist
    let participantArchiveSheet = responseSs.getSheetByName(ARCHIVE_PARTICIPANTS_SHEET_NAME);
    if (!participantArchiveSheet) {
      participantArchiveSheet = responseSs.insertSheet(ARCHIVE_PARTICIPANTS_SHEET_NAME);
      logMessage(`Deelnemers archief sheet '${ARCHIVE_PARTICIPANTS_SHEET_NAME}' aangemaakt.`);
    }
    
    // Track standard headers for consistency across different response sheets
    let standardHeaders = null;
    
    [OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME].forEach(sheetName => {
      const responseSheet = responseSs.getSheetByName(sheetName);
      if (!responseSheet) {
        logMessage(`WAARSCHUWING: Respons-sheet '${sheetName}' niet gevonden voor archivering.`);
        return;
      }
      
      const responseData = responseSheet.getDataRange().getValues();
      if (responseData.length < 2) return; // Only headers
      
      const responseHeaders = responseData[0];
      const eventColIdx = responseHeaders.indexOf(FORM_EVENT_QUESTION_TITLE);
      
      if (eventColIdx === -1) {
        logMessage(`WAARSCHUWING: Kolom '${FORM_EVENT_QUESTION_TITLE}' ontbreekt in respons-sheet '${sheetName}'. Kan deze sheet niet archiveren.`);
        return;
      }
      
      // Set standard headers on first iteration to ensure consistency
      if (!standardHeaders) {
        standardHeaders = responseHeaders.concat(['Bron Sheet']); // Add 'Bron Sheet' as last column
        
        // Set/verify headers in archive sheet
        if (participantArchiveSheet.getLastRow() === 0) {
          participantArchiveSheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
          logMessage(`Archief headers ingesteld: ${standardHeaders.join(', ')}`);
        } else {
          // Verify existing headers match expected structure
          const existingHeaders = participantArchiveSheet.getRange(1, 1, 1, participantArchiveSheet.getLastColumn()).getValues()[0];
          if (JSON.stringify(existingHeaders) !== JSON.stringify(standardHeaders)) {
            logMessage(`WAARSCHUWING: Bestaande headers in archief sheet komen niet overeen met verwachte headers voor '${sheetName}'.`);
            logMessage(`Verwacht: ${standardHeaders.join(', ')}`);
            logMessage(`Gevonden: ${existingHeaders.join(', ')}`);
          }
        }
      } else {
        // Verify current response sheet headers match the standard (excluding the 'Bron Sheet' column)
        const expectedResponseHeaders = standardHeaders.slice(0, -1); // Remove 'Bron Sheet' from comparison
        if (JSON.stringify(responseHeaders) !== JSON.stringify(expectedResponseHeaders)) {
          logMessage(`WAARSCHUWING: Headers in '${sheetName}' komen niet overeen met standaard headers.`);
          logMessage(`Standaard: ${expectedResponseHeaders.join(', ')}`);
          logMessage(`'${sheetName}': ${responseHeaders.join(', ')}`);
        }
      }
      
      const participantsToArchive = [];
      const rowsToStrikeThrough = [];
      
      // Find participants from archived clinics
      for (let i = 1; i < responseData.length; i++) {
        const row = responseData[i];
        const participantClinicName = (row[eventColIdx] || '').replace(/\s\(.*\)$/, '').trim();
        
        if (archivedClinicNames.has(participantClinicName)) {
          // Create archive row with original data + source sheet as last column
          // Format: [Timestamp, Email, First Name, Last Name, Event, Phone, DOB, City, Participant#, etc., Bron Sheet]
          const archiveRow = [...row, sheetName]; // Append source sheet as last column
          
          // Validate row length matches expected structure
          if (archiveRow.length !== standardHeaders.length) {
            logMessage(`WAARSCHUWING: Rij lengte (${archiveRow.length}) komt niet overeen met verwachte header lengte (${standardHeaders.length}) voor deelnemer in '${sheetName}'.`);
          }
          
          participantsToArchive.push(archiveRow);
          rowsToStrikeThrough.push(i + 1); // +1 because sheet rows are 1-indexed
          
          // Log the participant being archived for debugging
          const participantEmail = row[responseHeaders.indexOf(FORM_EMAIL_QUESTION_TITLE)] || 'Onbekend email';
          logMessage(`Archivering deelnemer: ${participantEmail} van clinic "${participantClinicName}" uit '${sheetName}'`);
        }
      }
      
      // Archive participants to the archive sheet
      if (participantsToArchive.length > 0) {
        const startRow = participantArchiveSheet.getLastRow() + 1;
        participantArchiveSheet.getRange(startRow, 1, participantsToArchive.length, participantsToArchive[0].length)
          .setValues(participantsToArchive);
        
        // Apply strike-through formatting to original rows
        rowsToStrikeThrough.forEach(rowNum => {
          const range = responseSheet.getRange(rowNum, 1, 1, responseSheet.getLastColumn());
          range.setFontLine('line-through');
        });
        
        logMessage(`${participantsToArchive.length} deelnemers van gearchiveerde clinics verplaatst naar '${ARCHIVE_PARTICIPANTS_SHEET_NAME}' en doorgestreept in '${sheetName}'.`);
      }
    });

    const plural = numArchived === 1 ? 'clinic' : 'clinics';
    const successMessage = `${numArchived} ${plural} gearchiveerd.`;
    logMessage(successMessage);

  } catch (e) {
    const errorMessage = `FOUT tijdens archiveren: ${e.toString()}\n${e.stack}`;
    Logger.log(errorMessage);
    logMessage(errorMessage);
  } finally {
    logMessage(`----- EINDE ${logPrefix} -----`);
    flushLogs(); // Ensures logs are always saved
  }
}