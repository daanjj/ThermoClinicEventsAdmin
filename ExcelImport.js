// This file contains functions related to importing participant data from Excel files.

function getExcelFiles() {
  const filesList = [];
  try {
    const folder = DriveApp.getFolderById(EXCEL_IMPORT_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
    while (files.hasNext()) {
      const file = files.next();
      filesList.push({
        name: file.getName(),
        id: file.getId()
      });
    }
  } catch (e) {
    Logger.log(`Could not get Excel files from folder ID ${EXCEL_IMPORT_FOLDER_ID}. Error: ${e.toString()}`);
    logMessage(`FOUT bij ophalen Excel-bestanden uit map ID ${EXCEL_IMPORT_FOLDER_ID}: ${e.message}`);
  }
  return filesList;
}

function showExcelImportDialog() {
  const htmlTemplate = HtmlService.createTemplateFromFile('ExcelImportDialog');
  htmlTemplate.excelFiles = getExcelFiles();
  SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setWidth(500).setHeight(400), 'Selecteer importbestand');
}

/**
 * MODIFIED FUNCTION (v4)
 * Processes an Excel file to add or update participants. This version now creates the
 * main Event Folder (if needed) and stores its ID in the 'Data clinics' sheet.
 * @param {string} fileId The Google Drive ID of the Excel file to process.
 * @returns {string} A summary message of the import results.
 */
function processExcelFile(fileId) {
  let tempSheetId = null; // To store the ID of the temporary Google Sheet
  const logPrefix = `Excel Import (File ID: ${fileId})`;
  logMessage(`----- START ${logPrefix} -----`);

  try {
    const excelFile = DriveApp.getFileById(fileId);
    logMessage(`Start conversie van Excel-bestand "${excelFile.getName()}" naar Google Sheet.`);

    // Convert Excel file to a temporary Google Sheet
    const resource = {
      title: `[TEMP] ${excelFile.getName()} - ${new Date().getTime()}`,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: excelFile.getParents().next().getId() }] // Keep it in the same folder as the original Excel
    };
    const tempSheet = Drive.Files.insert(resource, excelFile.getBlob());
    tempSheetId = tempSheet.id;
    logMessage(`Tijdelijke Google Sheet aangemaakt met ID: ${tempSheetId}`);

    const tempSpreadsheet = SpreadsheetApp.openById(tempSheetId);
    const importSheet = tempSpreadsheet.getSheets()[0];
    const importData = importSheet.getDataRange().getValues(); // Raw values (dates as Date objects)
    const importDisplayData = importSheet.getDataRange().getDisplayValues(); // Display values (dates as strings)

    if (importData.length < 2) {
      throw new Error("Import mislukt. Het Excel-bestand is leeg of bevat alleen een koptekst.");
    }
    
    // --- DYNAMIC HEADER MAPPING FOR EXCEL ---
    const columnMappings = {
      date: ['datum'], time: ['tijd'], location: ['locatie'],
      firstName: ['voornaam'], lastName: ['achternaam'], email: ['email', 'communications email address'],
      phone: ['telefoonnummer', 'telefonnummer', 'telefoon'], dob: ['geboortedatum', 'geboorte datum'], city: ['woonplaats', 'plaats']
    };
    const requiredHeaders = ['date', 'time', 'location', 'firstName', 'lastName', 'email'];
    
    const excelHeaders = importData[0].map(h => String(h).trim().toLowerCase()); // Headers from the first row of Excel
    importData.shift(); // Remove headers from data
    importDisplayData.shift(); // Also remove header from display values

    const excelHeaderMap = {};
    excelHeaders.forEach((header, index) => { if (header) excelHeaderMap[header] = index; });

    const columnIndexMap = {}; // Map internal key to actual Excel column index
    for (const key in columnMappings) {
      columnIndexMap[key] = -1; // Initialize as not found
      for (const possibleName of columnMappings[key]) {
        if (excelHeaderMap[possibleName] !== undefined) {
          columnIndexMap[key] = excelHeaderMap[possibleName];
          break; // Found a match, move to next mapping key
        }
      }
    }
    
    const missingHeaders = requiredHeaders.filter(key => columnIndexMap[key] === -1);
    if (missingHeaders.length > 0) {
      const missingHeaderNames = missingHeaders.map(key => `"${columnMappings[key].join(' of ')}"`).join(', ');
      throw new Error(`Import mislukt. Verplichte kolommen niet gevonden: ${missingHeaderNames}.`);
    }

    // Helper to parse various date formats
    const parseDate = (rawDate) => {
      if (rawDate instanceof Date && !isNaN(rawDate)) return rawDate; // Already a valid Date object
      if (typeof rawDate === 'string' && rawDate.trim() !== '') {
        // Try common formats
        let d = new Date(rawDate); // JavaScript's default parser
        if (!isNaN(d)) return d;
        
        // Try DD-MM-YYYY or DD.MM.YYYY
        const parts = rawDate.split(/[-/.]/);
        if (parts.length === 3) {
          const p1 = parseInt(parts[0], 10), p2 = parseInt(parts[1], 10), p3 = parseInt(parts[2], 10);
          if (!isNaN(p1) && !isNaN(p2) && !isNaN(p3)) {
            // Assume DD-MM-YYYY or similar
            if (p1 <= 31 && p2 <= 12) return new Date(p3, p2 - 1, p1); 
            // If it could be MM-DD-YYYY, you'd add more logic. For Dutch context, DD-MM is more common.
          }
        }
      }
      return null; // Could not parse
    };
    
    // --- STEP 1: Determine the event and its type from the first row of Excel data ---
    // Use the first data row (after headers) to identify the clinic
    const firstRowData = importData[0]; 
    const firstRowDisplay = importDisplayData[0]; 

    const eventDateRaw = firstRowData[columnIndexMap.date];
    const eventDate = parseDate(eventDateRaw);
    const eventTime = String(firstRowDisplay[columnIndexMap.time]).trim(); // Use display value for time string
    const eventLocation = String(firstRowDisplay[columnIndexMap.location]).trim(); // Use display value for location string
    
    if (!eventDate || !eventTime || !eventLocation) {
        throw new Error("Kon de eventgegevens (datum, tijd, locatie) niet uit de eerste rij van het Excel-bestand halen. Zorg dat deze kolommen gevuld zijn.");
    }
    
    const eventNameFromExcel = `${getDutchDateString(eventDate)} ${eventTime}, ${eventLocation}`;
    logMessage(`Excel import gestart voor event: "${eventNameFromExcel}"`);

    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const dataClinicsSheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
    const allDataClinicsData = dataClinicsSheet.getDataRange().getValues();
    const dataClinicsHeaders = allDataClinicsData.shift(); // Headers from Data Clinics sheet
    
    let clinicType = '';
    let clinicDataRowIndex = -1; // 1-based index in the Data Clinics sheet

    for (let i = 0; i < allDataClinicsData.length; i++) {
        const clinicRow = allDataClinicsData[i];
        const sheetDateRaw = clinicRow[DATE_COLUMN_INDEX - 1];
        if (!sheetDateRaw) continue;
        const sheetDate = new Date(sheetDateRaw);
        
        const reconstructedName = `${getDutchDateString(sheetDate)} ${String(clinicRow[TIME_COLUMN_INDEX - 1] || '').trim()}, ${String(clinicRow[LOCATION_COLUMN_INDEX - 1] || '').trim()}`;
        
        if (reconstructedName === eventNameFromExcel) {
            clinicType = String(clinicRow[TYPE_COLUMN_INDEX - 1] || '').trim().toLowerCase();
            clinicDataRowIndex = i + DATA_CLINICS_START_ROW; // 1-based row index in sheet
            break;
        }
    }

    if (!clinicType) {
        throw new Error(`Het event "${eventNameFromExcel}" uit het Excel-bestand kon niet worden gevonden in het '${DATA_CLINICS_SHEET_NAME}' tabblad. Zorg dat de datum, tijd en locatie exact overeenkomen.`);
    }

    // --- STEP 2: Pre-load existing participants and determine target sheet ---
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheetForAddsName = clinicType === 'open' ? OPEN_FORM_RESPONSE_SHEET_NAME : BESLOTEN_FORM_RESPONSE_SHEET_NAME;
    const targetSheetForAdds = ss.getSheetByName(targetSheetForAddsName);
    if (!targetSheetForAdds) throw new Error(`Kon het doeltabblad '${targetSheetForAddsName}' niet vinden.`);
    logMessage(`Event type: '${clinicType}'. Nieuwe deelnemers gaan naar: '${targetSheetForAddsName}'.`);

    // Build a map of existing participants by email + event name for quick lookup
    const existingParticipantsMap = {}; // Key: "email|eventname", Value: { sheetName: "...", rowIndex: X }
    [OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME].forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            logMessage(`WAARSCHUWING: Respons-sheet '${sheetName}' niet gevonden bij het controleren op bestaande deelnemers.`);
            return;
        }
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return; // Only headers
        const headers = data.shift();
        const headerMap = {};
        headers.forEach((h, i) => headerMap[h] = i);

        // Ensure required headers are present for lookup
        if (headerMap[FORM_EMAIL_QUESTION_TITLE] === undefined || headerMap[FORM_EVENT_QUESTION_TITLE] === undefined) {
            logMessage(`WAARSCHUWING: Kolom '${FORM_EMAIL_QUESTION_TITLE}' of '${FORM_EVENT_QUESTION_TITLE}' ontbreekt in respons-sheet '${sheetName}'. Kan bestaande deelnemers niet controleren.`);
            return;
        }

        data.forEach((row, index) => {
            const email = String(row[headerMap[FORM_EMAIL_QUESTION_TITLE]] || '').trim().toLowerCase();
            const eventName = String(row[headerMap[FORM_EVENT_QUESTION_TITLE]] || '').replace(/\s\(.*\)$/, '').trim();
            if (email && eventName) {
                const key = `${email}|${eventName}`;
                if (!existingParticipantsMap[key]) {
                    existingParticipantsMap[key] = { sheetName: sheetName, rowIndex: index + 2 }; // +2 for 1-based and header
                }
            }
        });
    });

    // --- STEP 3: Create or find the main Event Folder in Drive ---
    const eventDateFormattedForFolder = Utilities.formatDate(eventDate, FORMATTING_TIME_ZONE, DATE_FORMAT_YYYYMMDD);
    const eventTimeFormatted = eventTime.replace(/:|\.\|/g, ''); // Remove colons/dots for folder name
    const eventFolderName = `${eventDateFormattedForFolder} ${eventTimeFormatted} ${eventLocation}`;
    
    const parentFolder = DriveApp.getFolderById(PARENT_EVENT_FOLDER_ID);
    const eventFolders = parentFolder.getFoldersByName(eventFolderName);
    
    let eventFolder;
    if (eventFolders.hasNext()) {
      eventFolder = eventFolders.next();
      logMessage(`Bestaande Event Folder "${eventFolderName}" gevonden met ID: ${eventFolder.getId()}.`);
      
      // Check for duplicates
      if (eventFolders.hasNext()) {
        logMessage(`WAARSCHUWING: Meerdere mappen gevonden met naam "${eventFolderName}". Eerste map wordt gebruikt (ID: ${eventFolder.getId()}).`);
      }
    } else {
      eventFolder = parentFolder.createFolder(eventFolderName);
      logMessage(`Nieuwe Event Folder "${eventFolderName}" aangemaakt met ID: ${eventFolder.getId()}.`);
    }
    
    // Write Event Folder ID back to the Data Clinics sheet if it's missing
    const eventFolderIdColIdx = dataClinicsHeaders.indexOf(EVENT_FOLDER_ID_HEADER);
    if (eventFolderIdColIdx !== -1) {
      const currentEventFolderIdInSheet = dataClinicsSheet.getRange(clinicDataRowIndex, eventFolderIdColIdx + 1).getValue();
      if (!currentEventFolderIdInSheet) {
        dataClinicsSheet.getRange(clinicDataRowIndex, eventFolderIdColIdx + 1).setValue(eventFolder.getId());
        logMessage(`Event Folder ID ${eventFolder.getId()} opgeslagen voor clinic op rij ${clinicDataRowIndex} in '${DATA_CLINICS_SHEET_NAME}'.`);
      } else if (currentEventFolderIdInSheet !== eventFolder.getId()) {
        logMessage(`WAARSCHUWING: Event Folder ID in sheet (${currentEventFolderIdInSheet}) komt niet overeen met gevonden/aangemaakte ID (${eventFolder.getId()}) voor clinic op rij ${clinicDataRowIndex}. Sheet ID is NIET overschreven.`);
      }
    } else {
        logMessage(`WAARSCHUWING: Kolom '${EVENT_FOLDER_ID_HEADER}' niet gevonden in '${DATA_CLINICS_SHEET_NAME}'. Kan Event Folder ID niet opslaan.`);
    }


    // --- STEP 4: Process all participants from the Excel file ---
    let addCount = 0, updateCount = 0, errorCount = 0;
    const errorMessages = [];
    
    const targetHeadersForAdds = targetSheetForAdds.getRange(1, 1, 1, targetSheetForAdds.getLastColumn()).getValues()[0];
    const targetHeaderMapForAdds = {};
    targetHeadersForAdds.forEach((h, i) => targetHeaderMapForAdds[h] = i);

    for (let i = 0; i < importData.length; i++) {
      const row = importData[i]; // Raw data row (for Date objects)
      const displayRow = importDisplayData[i]; // Display data row (for formatted strings)
      const rowNumInExcel = i + 2; // Original row number in Excel file (1-based, +1 for headers)

      const email = displayRow[columnIndexMap.email] ? String(displayRow[columnIndexMap.email]).trim() : '';
      if (!email) {
          errorMessages.push(`Rij ${rowNumInExcel}: Overgeslagen, geen e-mailadres gevonden.`); 
          errorCount++; 
          continue;
      }

      const lookupKey = `${email.toLowerCase()}|${eventNameFromExcel}`;
      const existingParticipantInfo = existingParticipantsMap[lookupKey];
      
      const firstName = displayRow[columnIndexMap.firstName] ? String(displayRow[columnIndexMap.firstName]).trim() : '';
      const lastName = displayRow[columnIndexMap.lastName] ? String(displayRow[columnIndexMap.lastName]).trim() : '';
      const phone = columnIndexMap.phone !== -1 ? (displayRow[columnIndexMap.phone] ? String(displayRow[columnIndexMap.phone]).trim() : '') : '';
      const dobRaw = columnIndexMap.dob !== -1 ? row[columnIndexMap.dob] : null; // Use raw data for date object
      const city = columnIndexMap.city !== -1 ? (displayRow[columnIndexMap.city] ? String(displayRow[columnIndexMap.city]).trim() : '') : ''; 
      let dob = dobRaw ? parseDate(dobRaw) : null;
      
      if (existingParticipantInfo) {
        // Participant already exists, update their details
        const updateSheet = ss.getSheetByName(existingParticipantInfo.sheetName);
        if (!updateSheet) { // Should not happen if map was built correctly, but safety check
            errorMessages.push(`Rij ${rowNumInExcel} (${email}): Kon het tabblad '${existingParticipantInfo.sheetName}' niet vinden om deelnemer bij te werken.`);
            errorCount++;
            continue;
        }

        const updateSheetHeaders = updateSheet.getRange(1, 1, 1, updateSheet.getLastColumn()).getValues()[0];
        const updateHeaderMap = {};
        updateSheetHeaders.forEach((h, i) => updateHeaderMap[h] = i);
        
        logMessage(`Deelnemer ${email} bestaat al in '${existingParticipantInfo.sheetName}' op rij ${existingParticipantInfo.rowIndex}. Bijwerken...`);
        
        // Update fields if new data is present and columns exist
        if (firstName && updateHeaderMap[FORM_FIRST_NAME_QUESTION_TITLE] !== undefined) updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_FIRST_NAME_QUESTION_TITLE] + 1).setValue(firstName);
        if (lastName && updateHeaderMap[FORM_LAST_NAME_QUESTION_TITLE] !== undefined) updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_LAST_NAME_QUESTION_TITLE] + 1).setValue(lastName);
        if (phone && updateHeaderMap[FORM_PHONE_QUESTION_TITLE] !== undefined) updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_PHONE_QUESTION_TITLE] + 1).setValue(`'${phone}`); // Prefix with ' to force text
        if (dob && updateHeaderMap[FORM_DOB_QUESTION_TITLE] !== undefined) updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_DOB_QUESTION_TITLE] + 1).setValue(Utilities.formatDate(dob, FORMATTING_TIME_ZONE, 'dd-MM-yyyy'));
        if (city && updateHeaderMap[FORM_CITY_QUESTION_TITLE] !== undefined) updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_CITY_QUESTION_TITLE] + 1).setValue(city);
        
        updateCount++;
      } else {
        // New participant, add to the sheet and create Drive folder
        const bookedSeatsRange = dataClinicsSheet.getRange(clinicDataRowIndex, BOOKED_SEATS_COLUMN_INDEX);
        const currentBookedSeats = bookedSeatsRange.getValue() || 0;
        const isNonParticipant = isNonParticipantEmail(email);
        
        let participantSequenceNumber;
        if (isNonParticipant) {
          // Non-participant (test/host account): do NOT increment booked seats count
          logMessage(`Non-participant account gedetecteerd (${email}). Deelnemertelling NIET verhoogd.`);
          participantSequenceNumber = 'xx';
        } else {
          // Regular participant: increment the count
          const newBookedSeats = currentBookedSeats + 1;
          bookedSeatsRange.setValue(newBookedSeats); // Update booked seats count in Data Clinics sheet
          participantSequenceNumber = Utilities.formatString('%02d', newBookedSeats);
        }
        const participantSubfolderName = `${participantSequenceNumber} ${firstName} ${lastName}`.replace(/\s+/g, ' ').trim();
        const participantSubfolder = eventFolder.createFolder(participantSubfolderName);

        const newRowValues = [];
        targetHeadersForAdds.forEach(header => {
          let value = '';
          switch (header) {
            case 'Timestamp': value = new Date(); break;
            case FORM_EVENT_QUESTION_TITLE: value = eventNameFromExcel; break;
            case FORM_FIRST_NAME_QUESTION_TITLE: value = firstName; break;
            case FORM_LAST_NAME_QUESTION_TITLE: value = lastName; break;
            case FORM_EMAIL_QUESTION_TITLE: value = email; break;
            case FORM_PHONE_QUESTION_TITLE: value = `'${phone}`; break; // Prefix with ' to force text
            case FORM_DOB_QUESTION_TITLE: value = dob ? Utilities.formatDate(dob, FORMATTING_TIME_ZONE, 'dd-MM-yyyy') : ''; break;
            case FORM_CITY_QUESTION_TITLE: value = city; break;
            case FORM_REG_METHOD_QUESTION_TITLE: value = 'Excel Import'; break;
            case DEELNEMERNUMMER_HEADER: value = participantSequenceNumber; break;
            case DRIVE_FOLDER_ID_HEADER: value = participantSubfolder.getId(); break;
            default: value = ''; // Default for unhandled headers
          }
          newRowValues.push(value);
        });
        targetSheetForAdds.appendRow(newRowValues);
        addCount++;
        logMessage(`Deelnemer ${firstName} ${lastName} (nr. ${participantSequenceNumber}, email: ${email}) toegevoegd via Excel import.`);
      }
    }
    
    // --- STEP 5: Finalize ---
    if (addCount > 0) {
      Logger.log(`Excel import added participants. Triggering calendar sync for row ${clinicDataRowIndex}.`);
      syncCalendarEventFromSheet(clinicDataRowIndex); // Update calendar event with new count
    }
    
    let summaryMessage = `Import voltooid voor event "${eventNameFromExcel}".\n\nToegevoegd: ${addCount}\nBijgewerkt: ${updateCount}\nMislukt: ${errorCount}`;
    if (errorCount > 0) {
      summaryMessage += '\n\nFoutmeldingen:\n' + errorMessages.join('\n');
    }
    logMessage(`----- EINDE ${logPrefix}. Toegevoegd: ${addCount}, Bijgewerkt: ${updateCount}, Mislukt: ${errorCount} -----`);
    return summaryMessage;

  } catch (e) {
    const errorMessage = `Er is een kritieke fout opgetreden bij Excel import: ${e.message}`;
    Logger.log(`${logPrefix} FAILED: ${e.toString()}\n${e.stack}`);
    logMessage(`${logPrefix} FAILED: ${e.message}`);
    return errorMessage;
  } finally {
    if (tempSheetId) {
      try {
        DriveApp.getFileById(tempSheetId).setTrashed(true); // Delete the temporary Google Sheet
        logMessage(`Tijdelijke Google Sheet ID ${tempSheetId} verwijderd.`);
      } catch (deleteError) {
        Logger.log(`Could not delete temporary sheet ID ${tempSheetId}. Error: ${deleteError.toString()}`);
        logMessage(`WAARSCHUWING: Kon tijdelijke sheet ID ${tempSheetId} niet verwijderen: ${deleteError.message}`);
      }
    }
    flushLogs();
  }
}