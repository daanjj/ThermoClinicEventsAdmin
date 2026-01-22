// This file contains functions responsible for archiving old clinic data
// and associated participant responses.

function runManualArchive() {
  archiveOldClinics(true);
}

function runDailyArchive() {
  archiveOldClinics(false);
}

/**
 * Menu function to ensure all strikethrough participants are properly archived.
 * This function:
 * 1. Asks if already-archived strikethrough participants should be deleted
 * 2. Loops through all rows in Open and Besloten response sheets
 * 3. For each strikethrough row, checks if it exists in 'ARCHIEF deelnemers'
 * 4. If not present, adds the participant to the archive
 * 5. If already present and user chose to delete, removes the row from the source sheet
 * 6. Shows a summary dialog at the end
 */
function archiveStrikethroughParticipants() {
  logMessage(`----- START Archiveer doorgestreepte deelnemers -----`);
  
  const ui = SpreadsheetApp.getUi();
  
  // Ask user if they want to delete strikethrough rows that are already archived
  // Default is NO (don't delete) - require explicit confirmation for delete
  let shouldDelete = false;
  
  const deleteResponse = ui.alert(
    'Doorgestreepte deelnemers verwijderen?',
    'Wil je doorgestreepte deelnemers die al in het archief staan VERWIJDEREN uit de bron-sheets?\n\n' +
    'Klik CANCEL / ANNULEREN om alleen te archiveren zonder te verwijderen (aanbevolen).\n' +
    'Klik OK om te verwijderen.',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (deleteResponse === ui.Button.OK) {
    // Extra confirmation for destructive action
    const confirmDelete = ui.alert(
      '⚠️ Bevestig verwijderen',
      'Weet je zeker dat je alle doorgestreepte rijen die al gearchiveerd zijn wilt VERWIJDEREN?\n\n' +
      'Dit kan niet ongedaan worden gemaakt!',
      ui.ButtonSet.YES_NO
    );
    shouldDelete = (confirmDelete === ui.Button.YES);
  }
  
  logMessage(`Gebruiker koos: ${shouldDelete ? 'WEL verwijderen' : 'NIET verwijderen'} van reeds gearchiveerde doorgestreepte deelnemers.`);
  
  try {
    const responseSs = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get or create participant archive sheet
    let participantArchiveSheet = responseSs.getSheetByName(ARCHIVE_PARTICIPANTS_SHEET_NAME);
    if (!participantArchiveSheet) {
      participantArchiveSheet = responseSs.insertSheet(ARCHIVE_PARTICIPANTS_SHEET_NAME);
      logMessage(`Deelnemers archief sheet '${ARCHIVE_PARTICIPANTS_SHEET_NAME}' aangemaakt.`);
    }
    
    // Build set of already archived participants (normalized email + clinic name as key)
    const alreadyArchivedParticipants = new Set();
    let archiveEmailIdx = -1;
    let archiveEventIdx = -1;
    
    if (participantArchiveSheet.getLastRow() > 1) {
      const archiveParticipantData = participantArchiveSheet.getDataRange().getValues();
      const archiveHeaders = archiveParticipantData[0];
      archiveEmailIdx = archiveHeaders.indexOf(FORM_EMAIL_QUESTION_TITLE);
      archiveEventIdx = archiveHeaders.indexOf(FORM_EVENT_QUESTION_TITLE);
      
      if (archiveEmailIdx !== -1 && archiveEventIdx !== -1) {
        for (let i = 1; i < archiveParticipantData.length; i++) {
          const email = String(archiveParticipantData[i][archiveEmailIdx] || '').trim().toLowerCase();
          const eventName = normalizeClinicName(String(archiveParticipantData[i][archiveEventIdx] || '').replace(/\s\(.*\)$/, ''));
          if (email && eventName) {
            alreadyArchivedParticipants.add(`${email}|${eventName}`);
          }
        }
      }
    }
    
    logMessage(`${alreadyArchivedParticipants.size} deelnemers al in archief.`);
    
    let totalAdded = 0;
    let totalAlreadyArchived = 0;
    let totalDeleted = 0;
    let standardHeaders = null;
    
    [OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME].forEach(sheetName => {
      const responseSheet = responseSs.getSheetByName(sheetName);
      if (!responseSheet) {
        logMessage(`WAARSCHUWING: Sheet '${sheetName}' niet gevonden.`);
        return;
      }
      
      const responseData = responseSheet.getDataRange().getValues();
      if (responseData.length < 2) return;
      
      const responseHeaders = responseData[0];
      const eventColIdx = responseHeaders.indexOf(FORM_EVENT_QUESTION_TITLE);
      const emailColIdx = responseHeaders.indexOf(FORM_EMAIL_QUESTION_TITLE);
      
      if (eventColIdx === -1 || emailColIdx === -1) {
        logMessage(`WAARSCHUWING: Vereiste kolommen ontbreken in '${sheetName}'.`);
        return;
      }
      
      // Set standard headers on first iteration
      if (!standardHeaders) {
        standardHeaders = responseHeaders.concat(['Bron Sheet']);
        if (participantArchiveSheet.getLastRow() === 0) {
          participantArchiveSheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
          logMessage(`Archief headers ingesteld.`);
        }
      }
      
      const participantsToArchive = [];
      const rowsToDelete = []; // Store row numbers to delete (in reverse order later)
      let sheetAlreadyArchived = 0;
      
      // Loop through all rows and check for strikethrough
      for (let i = 1; i < responseData.length; i++) {
        const rowNum = i + 1;
        const row = responseData[i];
        
        // Check if row has strikethrough formatting
        const fontLine = responseSheet.getRange(rowNum, 1).getFontLine();
        if (fontLine !== 'line-through') {
          continue; // Skip non-strikethrough rows
        }
        
        const rawClinicName = String(row[eventColIdx] || '').replace(/\s\(.*\)$/, '').trim();
        const normalizedClinicName = normalizeClinicName(rawClinicName);
        const email = String(row[emailColIdx] || '').trim().toLowerCase();
        
        const archiveKey = `${email}|${normalizedClinicName}`;
        
        if (alreadyArchivedParticipants.has(archiveKey)) {
          sheetAlreadyArchived++;
          if (shouldDelete) {
            rowsToDelete.push(rowNum);
            logMessage(`Markeren voor verwijdering: ${email} van "${rawClinicName}" (${sheetName}) - staat al in archief`);
          }
          continue; // Already in archive
        }
        
        // Add to archive
        const archiveRow = [...row, sheetName];
        participantsToArchive.push(archiveRow);
        alreadyArchivedParticipants.add(archiveKey); // Prevent duplicates within same run
        logMessage(`Toevoegen aan archief: ${email} van "${rawClinicName}" (${sheetName})`);
      }
      
      // Write to archive sheet
      if (participantsToArchive.length > 0) {
        const startRow = participantArchiveSheet.getLastRow() + 1;
        participantArchiveSheet.getRange(startRow, 1, participantsToArchive.length, participantsToArchive[0].length)
          .setValues(participantsToArchive);
        
        SpreadsheetApp.flush();
        
        totalAdded += participantsToArchive.length;
        logMessage(`${participantsToArchive.length} doorgestreepte deelnemers toegevoegd aan archief vanuit '${sheetName}'.`);
      }
      
      // Delete rows if user chose to do so (delete from bottom to top to preserve row numbers)
      if (rowsToDelete.length > 0) {
        rowsToDelete.sort((a, b) => b - a); // Sort descending
        rowsToDelete.forEach(rowNum => {
          responseSheet.deleteRow(rowNum);
        });
        totalDeleted += rowsToDelete.length;
        logMessage(`${rowsToDelete.length} doorgestreepte rijen verwijderd uit '${sheetName}'.`);
      }
      
      totalAlreadyArchived += sheetAlreadyArchived;
      if (sheetAlreadyArchived > 0) {
        logMessage(`${sheetAlreadyArchived} doorgestreepte deelnemers in '${sheetName}' stonden al in archief.`);
      }
    });
    
    let message = `Archivering voltooid!\n\n` +
                  `${totalAdded} doorgestreepte deelnemers toegevoegd aan '${ARCHIVE_PARTICIPANTS_SHEET_NAME}'.\n` +
                  `${totalAlreadyArchived} doorgestreepte deelnemers stonden al in archief.`;
    
    if (shouldDelete && totalDeleted > 0) {
      message += `\n${totalDeleted} doorgestreepte rijen verwijderd uit bron-sheets.`;
    }
    
    logMessage(message.replace(/\n/g, ' '));
    ui.alert('Archivering doorgestreepte deelnemers', message, ui.ButtonSet.OK);
    
  } catch (e) {
    const errorMessage = `FOUT tijdens archiveren doorgestreepte deelnemers: ${e.toString()}\n${e.stack}`;
    Logger.log(errorMessage);
    logMessage(errorMessage);
    ui.alert('Fout', `Er is een fout opgetreden: ${e.message}`, ui.ButtonSet.OK);
  } finally {
    logMessage(`----- EINDE Archiveer doorgestreepte deelnemers -----`);
    flushLogs();
  }
}

/**
 * Utility function to retroactively fix participants that were missed during previous archiving runs.
 * This will:
 * 1. Find all participants in Open/Besloten sheets whose clinics are already in the archive
 * 2. Copy missing participants to the archive sheet
 * 3. Apply strike-through formatting to those rows
 * Run this manually from the Apps Script editor to fix historical data.
 */
function fixMissedArchivedParticipants() {
  logMessage(`----- START Herstel gemiste gearchiveerde deelnemers  -----`);
  
  try {
    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const responseSs = SpreadsheetApp.getActiveSpreadsheet();
    
    // Define the threshold for archiving (30 days old)
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    thirtyDaysAgo.setHours(0, 0, 0, 0);
    
    // Get archived clinic names from the archive sheet
    const archiveSheet = dataClinicsSpreadsheet.getSheetByName(ARCHIVE_SHEET_NAME);
    if (!archiveSheet) {
      logMessage(`Archief sheet '${ARCHIVE_SHEET_NAME}' niet gevonden. Niets te herstellen.`);
      SpreadsheetApp.getUi().alert(`Archief sheet '${ARCHIVE_SHEET_NAME}' niet gevonden.`);
      return;
    }
    
    const archiveData = archiveSheet.getDataRange().getValues();
    if (archiveData.length < 2) {
      logMessage(`Geen gearchiveerde clinics gevonden.`);
      SpreadsheetApp.getUi().alert(`Geen gearchiveerde clinics gevonden.`);
      return;
    }
    
    // Build set of archived clinic names (normalized for robust matching)
    const archivedClinicNames = new Set();
    const normalizedArchivedClinicNames = new Set();
    for (let i = 1; i < archiveData.length; i++) {
      const row = archiveData[i];
      const dateValue = row[DATE_COLUMN_INDEX - 1];
      if (!dateValue) continue;
      const clinicDate = new Date(dateValue);
      if (isNaN(clinicDate.getTime())) continue;
      
      const clinicName = `${getDutchDateString(clinicDate)} ${String(row[TIME_COLUMN_INDEX - 1] || '').trim()}, ${String(row[LOCATION_COLUMN_INDEX - 1] || '').trim()}`;
      archivedClinicNames.add(clinicName);
      normalizedArchivedClinicNames.add(normalizeClinicName(clinicName));
    }
    
    logMessage(`${archivedClinicNames.size} gearchiveerde clinics gevonden.`);
    logMessage(`Gearchiveerde clinic namen (genormaliseerd): ${Array.from(normalizedArchivedClinicNames).slice(0, 5).join('; ')}${normalizedArchivedClinicNames.size > 5 ? '...' : ''}`);
    
    // Get or create participant archive sheet
    let participantArchiveSheet = responseSs.getSheetByName(ARCHIVE_PARTICIPANTS_SHEET_NAME);
    if (!participantArchiveSheet) {
      participantArchiveSheet = responseSs.insertSheet(ARCHIVE_PARTICIPANTS_SHEET_NAME);
      logMessage(`Deelnemers archief sheet '${ARCHIVE_PARTICIPANTS_SHEET_NAME}' aangemaakt.`);
    }
    
    // Build set of already archived participants (email + normalized clinic name as key)
    const alreadyArchivedParticipants = new Set();
    if (participantArchiveSheet.getLastRow() > 1) {
      const archiveParticipantData = participantArchiveSheet.getDataRange().getValues();
      const archiveHeaders = archiveParticipantData[0];
      const archiveEmailIdx = archiveHeaders.indexOf(FORM_EMAIL_QUESTION_TITLE);
      const archiveEventIdx = archiveHeaders.indexOf(FORM_EVENT_QUESTION_TITLE);
      
      if (archiveEmailIdx !== -1 && archiveEventIdx !== -1) {
        for (let i = 1; i < archiveParticipantData.length; i++) {
          const email = String(archiveParticipantData[i][archiveEmailIdx] || '').trim().toLowerCase();
          const eventName = normalizeClinicName(String(archiveParticipantData[i][archiveEventIdx] || '').replace(/\s\(.*\)$/, ''));
          if (email && eventName) {
            alreadyArchivedParticipants.add(`${email}|${eventName}`);
          }
        }
      }
    }
    
    logMessage(`${alreadyArchivedParticipants.size} deelnemers al in archief.`);
    
    let totalFixed = 0;
    let totalStrikethrough = 0;
    let standardHeaders = null;
    
    [OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME].forEach(sheetName => {
      const responseSheet = responseSs.getSheetByName(sheetName);
      if (!responseSheet) {
        logMessage(`WAARSCHUWING: Sheet '${sheetName}' niet gevonden.`);
        return;
      }
      
      const responseData = responseSheet.getDataRange().getValues();
      if (responseData.length < 2) return;
      
      const responseHeaders = responseData[0];
      const eventColIdx = responseHeaders.indexOf(FORM_EVENT_QUESTION_TITLE);
      const emailColIdx = responseHeaders.indexOf(FORM_EMAIL_QUESTION_TITLE);
      
      if (eventColIdx === -1 || emailColIdx === -1) {
        logMessage(`WAARSCHUWING: Vereiste kolommen ontbreken in '${sheetName}'.`);
        return;
      }
      
      // Set standard headers on first iteration
      if (!standardHeaders) {
        standardHeaders = responseHeaders.concat(['Bron Sheet']);
        if (participantArchiveSheet.getLastRow() === 0) {
          participantArchiveSheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
        }
      }
      
      const participantsToArchive = [];
      const rowsToStrikeThrough = [];
      
      for (let i = 1; i < responseData.length; i++) {
        const row = responseData[i];
        const rowNum = i + 1;
        
        // Check if already has strike-through
        const currentFormat = responseSheet.getRange(rowNum, 1).getFontLine();
        const hasStrikethrough = currentFormat === 'line-through';
        
        const rawParticipantClinicName = String(row[eventColIdx] || '').replace(/\s\(.*\)$/, '').trim();
        const normalizedParticipantClinicName = normalizeClinicName(rawParticipantClinicName);
        const participantEmail = String(row[emailColIdx] || '').trim().toLowerCase();
        
        // Check if this participant should be archived:
        // 1. Their clinic is in the archive (using normalized matching), OR
        // 2. Their clinic date (parsed from the clinic name) is older than 30 days
        let shouldArchive = normalizedArchivedClinicNames.has(normalizedParticipantClinicName);
        
        if (!shouldArchive && rawParticipantClinicName) {
          // Try to parse date from clinic name (format: "dag dd maand yyyy HH:MM-HH:MM, Locatie")
          const clinicDate = parseDutchDateFromClinicName(rawParticipantClinicName);
          if (clinicDate && clinicDate < thirtyDaysAgo) {
            shouldArchive = true;
            logMessage(`Clinic "${rawParticipantClinicName}" is ouder dan 30 dagen (datum: ${clinicDate.toLocaleDateString('nl-NL')})`);
          }
        }
        
        if (shouldArchive) {
          const archiveKey = `${participantEmail}|${normalizedParticipantClinicName}`;
          
          // Add to archive if not already there
          if (!alreadyArchivedParticipants.has(archiveKey)) {
            const archiveRow = [...row, sheetName];
            participantsToArchive.push(archiveRow);
            alreadyArchivedParticipants.add(archiveKey); // Prevent duplicates within same run
            logMessage(`Toevoegen aan archief: ${participantEmail} van "${rawParticipantClinicName}" (${sheetName})`);
          }
          
          // Apply strike-through if not already applied
          if (!hasStrikethrough) {
            rowsToStrikeThrough.push(rowNum);
          }
        }
      }
      
      // Archive missing participants - with verification before applying strike-through
      if (participantsToArchive.length > 0) {
        const startRow = participantArchiveSheet.getLastRow() + 1;
        const expectedEndRow = startRow + participantsToArchive.length - 1;
        
        participantArchiveSheet.getRange(startRow, 1, participantsToArchive.length, participantsToArchive[0].length)
          .setValues(participantsToArchive);
        
        // Force flush to ensure data is written
        SpreadsheetApp.flush();
        
        // VERIFICATION: Read back to confirm data was actually written
        const actualLastRow = participantArchiveSheet.getLastRow();
        const archiveSuccessful = actualLastRow >= expectedEndRow;
        
        if (!archiveSuccessful) {
          logMessage(`KRITIEKE FOUT: Archivering mislukt voor ${participantsToArchive.length} deelnemers uit '${sheetName}'. ` +
                     `Verwacht einde rij: ${expectedEndRow}, werkelijk: ${actualLastRow}. ` +
                     `Strike-through wordt NIET toegepast om dataverlies te voorkomen.`);
          return; // Skip strike-through for this sheet
        }
        
        totalFixed += participantsToArchive.length;
        logMessage(`${participantsToArchive.length} ontbrekende deelnemers toegevoegd aan archief vanuit '${sheetName}'.`);
        
        // Only apply strike-through after successful archive verification
        if (rowsToStrikeThrough.length > 0) {
          rowsToStrikeThrough.forEach(rowNum => {
            responseSheet.getRange(rowNum, 1, 1, responseSheet.getLastColumn()).setFontLine('line-through');
          });
          totalStrikethrough += rowsToStrikeThrough.length;
          logMessage(`${rowsToStrikeThrough.length} rijen doorgestreept in '${sheetName}'.`);
        }
      } else if (rowsToStrikeThrough.length > 0) {
        // No new participants to archive but some rows need strike-through 
        // (already in archive but not struck through)
        rowsToStrikeThrough.forEach(rowNum => {
          responseSheet.getRange(rowNum, 1, 1, responseSheet.getLastColumn()).setFontLine('line-through');
        });
        totalStrikethrough += rowsToStrikeThrough.length;
        logMessage(`${rowsToStrikeThrough.length} rijen doorgestreept in '${sheetName}' (deelnemers al in archief).`);
      }
    });
    
    const message = `Herstel voltooid!\n\n${totalFixed} deelnemers toegevoegd aan archief.\n${totalStrikethrough} rijen doorgestreept.`;
    logMessage(message);
    SpreadsheetApp.getUi().alert(message);
    
  } catch (e) {
    const errorMessage = `FOUT tijdens herstel: ${e.toString()}\n${e.stack}`;
    Logger.log(errorMessage);
    logMessage(errorMessage);
    SpreadsheetApp.getUi().alert(`Er is een fout opgetreden: ${e.message}`);
  } finally {
    logMessage(`----- EINDE Herstel gemiste gearchiveerde deelnemers -----`);
    flushLogs();
  }
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
    
    // Convert archivedClinicNames Set to normalized versions for more robust matching
    const normalizedArchivedClinicNames = new Set();
    archivedClinicNames.forEach(name => {
      normalizedArchivedClinicNames.add(normalizeClinicName(name));
    });
    
    logMessage(`Gearchiveerde clinic namen (genormaliseerd): ${Array.from(normalizedArchivedClinicNames).join('; ')}`);
    
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
        const rawParticipantClinicName = (row[eventColIdx] || '').replace(/\s\(.*\)$/, '').trim();
        const normalizedParticipantClinicName = normalizeClinicName(rawParticipantClinicName);
        
        if (normalizedArchivedClinicNames.has(normalizedParticipantClinicName)) {
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
          logMessage(`Archivering deelnemer: ${participantEmail} van clinic "${rawParticipantClinicName}" uit '${sheetName}'`);
        }
      }
      
      // Archive participants to the archive sheet - with verification before applying strike-through
      if (participantsToArchive.length > 0) {
        const startRow = participantArchiveSheet.getLastRow() + 1;
        const expectedEndRow = startRow + participantsToArchive.length - 1;
        
        // Write to archive sheet
        participantArchiveSheet.getRange(startRow, 1, participantsToArchive.length, participantsToArchive[0].length)
          .setValues(participantsToArchive);
        
        // Force flush to ensure data is written
        SpreadsheetApp.flush();
        
        // VERIFICATION: Read back to confirm data was actually written
        const actualLastRow = participantArchiveSheet.getLastRow();
        const archiveSuccessful = actualLastRow >= expectedEndRow;
        
        if (!archiveSuccessful) {
          logMessage(`KRITIEKE FOUT: Archivering mislukt voor ${participantsToArchive.length} deelnemers uit '${sheetName}'. ` +
                     `Verwacht einde rij: ${expectedEndRow}, werkelijk: ${actualLastRow}. ` +
                     `Strike-through wordt NIET toegepast om dataverlies te voorkomen.`);
          return; // Skip strike-through for this sheet
        }
        
        // Additional verification: check that the first archived row contains expected data
        const verificationData = participantArchiveSheet.getRange(startRow, 1, 1, participantsToArchive[0].length).getValues()[0];
        const firstArchivedEmail = participantsToArchive[0][responseHeaders.indexOf(FORM_EMAIL_QUESTION_TITLE)] || '';
        const verificationEmail = verificationData[responseHeaders.indexOf(FORM_EMAIL_QUESTION_TITLE)] || '';
        
        if (String(firstArchivedEmail).trim().toLowerCase() !== String(verificationEmail).trim().toLowerCase()) {
          logMessage(`KRITIEKE FOUT: Verificatie mislukt - gearchiveerde data komt niet overeen. ` +
                     `Verwacht email: "${firstArchivedEmail}", gevonden: "${verificationEmail}". ` +
                     `Strike-through wordt NIET toegepast.`);
          return; // Skip strike-through for this sheet
        }
        
        logMessage(`Archivering geverifieerd: ${participantsToArchive.length} deelnemers succesvol geschreven naar rij ${startRow}-${expectedEndRow}.`);
        
        // Only apply strike-through after successful verification
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

/**
 * Helper function to parse a Dutch date from a clinic name string.
 * Expected format: "dag dd maand yyyy HH:MM-HH:MM, Locatie"
 * Example: "zaterdag 7 december 2025 10:00-13:00, Amsterdam"
 * @param {string} clinicName - The clinic name to parse
 * @returns {Date|null} - The parsed date or null if parsing fails
 */
function parseDutchDateFromClinicName(clinicName) {
  if (!clinicName) return null;
  
  const monthNamesDutch = {
    'januari': 0, 'februari': 1, 'maart': 2, 'april': 3, 'mei': 4, 'juni': 5,
    'juli': 6, 'augustus': 7, 'september': 8, 'oktober': 9, 'november': 10, 'december': 11
  };
  
  try {
    // Pattern: day-name day month year (e.g., "zaterdag 7 december 2025")
    const pattern = /(\d{1,2})\s+(januari|februari|maart|april|mei|juni|juli|augustus|september|oktober|november|december)\s+(\d{4})/i;
    const match = clinicName.match(pattern);
    
    if (match) {
      const day = parseInt(match[1], 10);
      const month = monthNamesDutch[match[2].toLowerCase()];
      const year = parseInt(match[3], 10);
      
      if (!isNaN(day) && month !== undefined && !isNaN(year)) {
        const date = new Date(year, month, day);
        if (!isNaN(date.getTime())) {
          return date;
        }
      }
    }
    return null;
  } catch (e) {
    Logger.log(`Error parsing date from clinic name "${clinicName}": ${e.message}`);
    return null;
  }
}

/**
 * Helper function to normalize a clinic name for robust comparison.
 * This handles common string matching issues like:
 * - Multiple spaces
 * - Leading/trailing whitespace
 * - Inconsistent capitalization
 * - Non-breaking spaces or other unicode whitespace
 * 
 * @param {string} clinicName - The clinic name to normalize
 * @returns {string} - The normalized clinic name for comparison
 */
function normalizeClinicName(clinicName) {
  if (!clinicName) return '';
  
  return String(clinicName)
    .toLowerCase()                           // Normalize case
    .replace(/\s+/g, ' ')                    // Collapse multiple whitespace to single space
    .replace(/[\u00A0\u2007\u202F]/g, ' ')   // Replace non-breaking spaces with regular space
    .replace(/\s*,\s*/g, ', ')               // Normalize comma spacing
    .replace(/\s*-\s*/g, '-')                // Normalize dash spacing  
    .trim();                                  // Remove leading/trailing whitespace
}