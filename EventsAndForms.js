// This file handles logic related to clinic events (time changes, calendar sync)
// and populating form dropdowns based on clinic data.

/**
 * Triggered on an edit in 'Data clinics'. If the date, time, or location column is changed,
 * this function updates corresponding entries in response sheets AND renames the event folder in Drive.
 * @param {Object} e The event object from the onEdit trigger.
 */
function handleEventChange(e) {
  // --- Step 1: Validate the edit ---
  if (!e || !e.range || !e.oldValue || !e.value || e.oldValue === e.value) {
    return; // Not a relevant value change
  }
  
  const sheet = e.range.getSheet();
  if (sheet.getName() !== DATA_CLINICS_SHEET_NAME) {
    return;
  }
  
  const editedRow = e.range.getRow();
  if (editedRow < DATA_CLINICS_START_ROW) {
    return; // Ignore header edits
  }
  
  const editedColumn = e.range.getColumn();
  
  // Check if the edited column is date, time, or location
  if (editedColumn !== DATE_COLUMN_INDEX && editedColumn !== TIME_COLUMN_INDEX && editedColumn !== LOCATION_COLUMN_INDEX) {
    return; // Not a relevant column
  }
  
  const columnNames = ['Datum', 'Tijdstip', 'Locatie'];
  const columnName = columnNames[editedColumn - 1] || 'Onbekend';
  logMessage(`${columnName} wijziging gedetecteerd in '${DATA_CLINICS_SHEET_NAME}' op rij ${editedRow}. Oud: "${e.oldValue}", Nieuw: "${e.value}".`);

  // --- Step 2: Read row data and construct event names ---
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const eventFolderIdColIdx = headers.indexOf(EVENT_FOLDER_ID_HEADER);
  
  const rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateValue = rowData[DATE_COLUMN_INDEX - 1];
  const timeValue = rowData[TIME_COLUMN_INDEX - 1];
  const locationValue = rowData[LOCATION_COLUMN_INDEX - 1];
  const eventFolderId = (eventFolderIdColIdx !== -1) ? rowData[eventFolderIdColIdx] : null;

  if (!dateValue || !timeValue || !locationValue) {
    logMessage(`FOUT: Kan event niet identificeren. Datum, tijd of locatie ontbreekt op rij ${editedRow}.`);
    return;
  }

  // --- Step 3: Rename the Google Drive Folder ---
  if (eventFolderId) {
    try {
      const dateFormatted = Utilities.formatDate(new Date(dateValue), FORMATTING_TIME_ZONE, DATE_FORMAT_YYYYMMDD);
      const timeFormatted = String(timeValue).trim().replace(/:|\./g, '');
      const newFolderName = `${dateFormatted} ${timeFormatted} ${locationValue}`;
      const folder = DriveApp.getFolderById(eventFolderId);
      const currentFolderName = folder.getName();
      
      if (currentFolderName !== newFolderName) {
        folder.setName(newFolderName);
        logMessage(`Event folder hernoemd van "${currentFolderName}" naar "${newFolderName}".`);
      } else {
        logMessage(`Event folder naam is al correct: "${newFolderName}".`);
      }
    } catch (driveError) {
      logMessage(`WAARSCHUWING: Kon de event folder met ID "${eventFolderId}" niet hernoemen. Fout: ${driveError.message}`);
      Logger.log(`Could not rename folder with ID ${eventFolderId}. Error: ${driveError.toString()}`);
    }
  } else {
    logMessage(`Info: Geen Event Folder ID gevonden op rij ${editedRow}. Hernoemen overgeslagen.`);
  }

  // --- Step 4: Update participant entries in response sheets ---
  const dateString = getDutchDateString(new Date(dateValue));
  
  // Construct old and new event names based on which column was changed
  let oldEventNameBase, newEventNameBase;
  
  if (editedColumn === DATE_COLUMN_INDEX) {
    const oldDateString = getDutchDateString(new Date(e.oldValue));
    oldEventNameBase = `${oldDateString} ${String(timeValue).trim()}, ${String(locationValue).trim()}`;
    newEventNameBase = `${dateString} ${String(timeValue).trim()}, ${String(locationValue).trim()}`;
  } else if (editedColumn === TIME_COLUMN_INDEX) {
    oldEventNameBase = `${dateString} ${String(e.oldValue).trim()}, ${String(locationValue).trim()}`;
    newEventNameBase = `${dateString} ${String(timeValue).trim()}, ${String(locationValue).trim()}`;
  } else if (editedColumn === LOCATION_COLUMN_INDEX) {
    oldEventNameBase = `${dateString} ${String(timeValue).trim()}, ${String(e.oldValue).trim()}`;
    newEventNameBase = `${dateString} ${String(timeValue).trim()}, ${String(locationValue).trim()}`;
  }
  
  logMessage(`Zoeken naar deelnemers voor event: "${oldEventNameBase}" om bij te werken naar "${newEventNameBase}".`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToUpdate = [OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME];
  let totalUpdates = 0;

  sheetsToUpdate.forEach(sheetName => {
    const responseSheet = ss.getSheetByName(sheetName);
    if (!responseSheet) return;

    const dataRange = responseSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values.shift();
    const eventColIdx = headers.indexOf(FORM_EVENT_QUESTION_TITLE);

    if (eventColIdx === -1) return;
    
    let updatesInSheet = 0;
    const newValues = values.map(row => {
      const originalValue = row[eventColIdx] || '';
      const eventNameInSheet = originalValue.replace(/\s\(.*\)$/, '').trim();

      if (eventNameInSheet === oldEventNameBase) {
        const seatCountMatch = originalValue.match(/\s\(.*\)$/);
        const seatCountSuffix = seatCountMatch ? seatCountMatch[0] : '';
        row[eventColIdx] = newEventNameBase + seatCountSuffix;
        updatesInSheet++;
      }
      return row;
    });

    if (updatesInSheet > 0) {
      responseSheet.getRange(2, 1, newValues.length, newValues[0].length).setValues(newValues);
      logMessage(`${updatesInSheet} deelnemer(s) bijgewerkt in tabblad '${sheetName}'.`);
      totalUpdates += updatesInSheet;
    }
  });

  if (totalUpdates > 0) {
    logMessage(`Totaal ${totalUpdates} deelnemersrecords bijgewerkt. Formulieren worden nu gesynchroniseerd.`);
  } else {
    logMessage(`Geen overeenkomende deelnemers gevonden om bij te werken.`);
  }
}

/**
 * Handles clinic type changes (Open <-> Besloten) by moving all participants 
 * from one response sheet to the other.
 * @param {Object} e The event object from the onEdit trigger.
 */
function handleClinicTypeChange(e) {
  // --- Step 1: Validate the edit ---
  if (!e || !e.range || !e.oldValue || !e.value || e.oldValue === e.value) {
    return; // Not a relevant value change
  }
  
  const sheet = e.range.getSheet();
  if (sheet.getName() !== DATA_CLINICS_SHEET_NAME) {
    return;
  }
  
  const editedRow = e.range.getRow();
  if (editedRow < DATA_CLINICS_START_ROW) {
    return; // Ignore header edits
  }
  
  const editedColumn = e.range.getColumn();
  
  // Check if the edited column is the Type column
  if (editedColumn !== TYPE_COLUMN_INDEX) {
    return; // Not the type column
  }
  
  const oldType = String(e.oldValue).trim().toLowerCase();
  const newType = String(e.value).trim().toLowerCase();
  
  // Validate types
  if (!['open', 'besloten'].includes(oldType) || !['open', 'besloten'].includes(newType)) {
    logMessage(`WAARSCHUWING: Onverwachte type waarde. Oud: "${e.oldValue}", Nieuw: "${e.value}". Verwacht 'Open' of 'Besloten'.`);
    return;
  }
  
  logMessage(`Type wijziging gedetecteerd in '${DATA_CLINICS_SHEET_NAME}' op rij ${editedRow}. Oud: "${e.oldValue}", Nieuw: "${e.value}".`);
  
  // --- Step 2: Get event details ---
  const rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateValue = rowData[DATE_COLUMN_INDEX - 1];
  const timeValue = rowData[TIME_COLUMN_INDEX - 1];
  const locationValue = rowData[LOCATION_COLUMN_INDEX - 1];
  
  if (!dateValue || !timeValue || !locationValue) {
    logMessage(`FOUT: Kan event niet identificeren. Datum, tijd of locatie ontbreekt op rij ${editedRow}.`);
    return;
  }
  
  const dateString = getDutchDateString(new Date(dateValue));
  const eventName = `${dateString} ${String(timeValue).trim()}, ${String(locationValue).trim()}`;
  
  logMessage(`Deelnemers verplaatsen voor event: "${eventName}" van ${oldType} naar ${newType}.`);
  
  // --- Step 3: Move participants from old sheet to new sheet ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSheetName = oldType === 'open' ? OPEN_FORM_RESPONSE_SHEET_NAME : BESLOTEN_FORM_RESPONSE_SHEET_NAME;
  const newSheetName = newType === 'open' ? OPEN_FORM_RESPONSE_SHEET_NAME : BESLOTEN_FORM_RESPONSE_SHEET_NAME;
  
  const oldSheet = ss.getSheetByName(oldSheetName);
  const newSheet = ss.getSheetByName(newSheetName);
  
  if (!oldSheet || !newSheet) {
    logMessage(`FOUT: Kon respons-tabbladen niet vinden. Oud: '${oldSheetName}', Nieuw: '${newSheetName}'.`);
    return;
  }
  
  // Get all data from old sheet
  const oldSheetData = oldSheet.getDataRange().getValues();
  if (oldSheetData.length < 2) {
    logMessage(`Geen deelnemers gevonden in '${oldSheetName}' om te verplaatsen.`);
    return;
  }
  
  const oldHeaders = oldSheetData[0];
  const eventColIdx = oldHeaders.indexOf(FORM_EVENT_QUESTION_TITLE);
  
  if (eventColIdx === -1) {
    logMessage(`FOUT: Kolom '${FORM_EVENT_QUESTION_TITLE}' niet gevonden in '${oldSheetName}'.`);
    return;
  }
  
  // Find rows that match this event
  const rowsToMove = [];
  const rowIndicesToDelete = [];
  
  for (let i = 1; i < oldSheetData.length; i++) {
    const row = oldSheetData[i];
    const eventNameInSheet = (row[eventColIdx] || '').replace(/\s\(.*\)$/, '').trim();
    
    if (eventNameInSheet === eventName) {
      rowsToMove.push(row);
      rowIndicesToDelete.push(i + 1); // +1 because sheet rows are 1-indexed
    }
  }
  
  if (rowsToMove.length === 0) {
    logMessage(`Geen deelnemers gevonden voor event "${eventName}" in '${oldSheetName}'.`);
    return;
  }
  
  logMessage(`${rowsToMove.length} deelnemer(s) gevonden om te verplaatsen van '${oldSheetName}' naar '${newSheetName}'.`);
  
  // --- Step 4: Verify headers match between sheets ---
  const newSheetData = newSheet.getDataRange().getValues();
  const newHeaders = newSheetData[0];
  
  // Check if headers are identical
  if (oldHeaders.length !== newHeaders.length || !oldHeaders.every((h, idx) => h === newHeaders[idx])) {
    logMessage(`WAARSCHUWING: Headers in '${oldSheetName}' en '${newSheetName}' komen niet overeen. Deelnemers kunnen niet worden verplaatst.`);
    Logger.log(`Old headers: ${JSON.stringify(oldHeaders)}`);
    Logger.log(`New headers: ${JSON.stringify(newHeaders)}`);
    return;
  }
  
  // --- Step 5: Copy rows to new sheet ---
  const lastRowInNewSheet = newSheet.getLastRow();
  newSheet.insertRowsAfter(lastRowInNewSheet, rowsToMove.length);
  
  rowsToMove.forEach((row, index) => {
    const targetRow = lastRowInNewSheet + index + 1;
    newSheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  });
  
  logMessage(`${rowsToMove.length} deelnemer(s) gekopieerd naar '${newSheetName}'.`);
  
  // --- Step 6: Delete rows from old sheet (in reverse order to avoid index shifting) ---
  rowIndicesToDelete.reverse().forEach(rowIndex => {
    oldSheet.deleteRow(rowIndex);
  });
  
  logMessage(`${rowsToMove.length} deelnemer(s) verwijderd uit '${oldSheetName}'.`);
  logMessage(`Type wijziging afgerond: ${rowsToMove.length} deelnemer(s) verplaatst van ${oldType} naar ${newType}.`);
}

function updateAllFormDropdowns() {
  populateFormDropdown('open');
  populateFormDropdown('besloten');
  populateCoreAppFormDropdown();
}

function syncCalendarEventFromSheet(rowNum) {
  if (rowNum < DATA_CLINICS_START_ROW) return;
  const sheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID).getSheetByName(DATA_CLINICS_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[String(header).trim()] = index;
  });
  const requiredHeaders = ['Datum', 'Tijdstip', 'Locatie', 'Aantal boekingen', CALENDAR_EVENT_ID_HEADER];
  for (const h of requiredHeaders) {
    if (headerMap[h] === undefined) {
      const errorMsg = `Required column "${h}" not found in '${DATA_CLINICS_SHEET_NAME}' sheet.` +
                       ` Please ensure header "${h}" exists.`; // Added more descriptive message
      Logger.log(errorMsg);
      logMessage(`FOUT: ${errorMsg}`);
      return;
    }
  }
  
  // Try to find max seats column by checking the correct header name
  let maxSeatsColIdx = -1;
  const possibleMaxSeatsHeaders = ['Maximum aantal deelnemers', 'Maximum aantal', 'Max seats', 'Maximaal aantal', 'Maximum', 'Max'];
  for (const headerName of possibleMaxSeatsHeaders) {
    if (headerMap[headerName] !== undefined) {
      maxSeatsColIdx = headerMap[headerName];
      break;
    }
  }
  
  // Fallback to hardcoded column index if header not found
  if (maxSeatsColIdx === -1) {
    maxSeatsColIdx = MAX_SEATS_COLUMN_INDEX - 1; // Convert to 0-based index
    Logger.log(`Warning: Max seats header not found in headerMap, using fallback column index ${MAX_SEATS_COLUMN_INDEX}`);
  }
  
  const dateValue = rowData[headerMap['Datum']];
  const timeValue = rowData[headerMap['Tijdstip']];
  const location = rowData[headerMap['Locatie']];
  const bookedSeatsRaw = rowData[headerMap['Aantal boekingen']];
  const maxSeatsRaw = rowData[maxSeatsColIdx]; // Get max seats using determined column index
  const eventId = rowData[headerMap[CALENDAR_EVENT_ID_HEADER]];

  const bookedSeats = parseInt(bookedSeatsRaw, 10) || 0;
  const maxSeats = parseInt(maxSeatsRaw, 10) || 0;
  
  // Debug logging
  Logger.log(`syncCalendarEventFromSheet row ${rowNum}: maxSeatsRaw="${maxSeatsRaw}", maxSeats=${maxSeats}, bookedSeats=${bookedSeats}`);
  Logger.log(`syncCalendarEventFromSheet row ${rowNum}: maxSeatsColIdx=${maxSeatsColIdx}, eventId="${eventId}"`);
  logMessage(`DEBUG: Sync calendar voor rij ${rowNum} - Max seats: ${maxSeats}, Booked seats: ${bookedSeats}, Event ID: ${eventId || 'geen'}`);
  
  // Check if max seats is 0 - if so, delete the calendar event
  if (maxSeats === 0) {
    logMessage(`Max seats = 0 gedetecteerd voor rij ${rowNum}. Verwijderen van agenda-item...`);
    Logger.log(`Max seats = 0 detected for row ${rowNum}, attempting to delete calendar event`);
    if (eventId) {
      try {
        const calendar = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);
        if (calendar) {
          const event = calendar.getEventById(eventId);
          if (event) {
            event.deleteEvent();
            // Clear the calendar event ID from the sheet
            sheet.getRange(rowNum, headerMap[CALENDAR_EVENT_ID_HEADER] + 1).setValue('');
            logMessage(`Agenda-item verwijderd voor event op rij ${rowNum} (max seats = 0). Locatie: "${location}".`);
            Logger.log(`Deleted calendar event for row ${rowNum} due to max seats = 0.`);
          }
        }
      } catch (e) {
        logMessage(`WAARSCHUWING: Kon agenda-item niet verwijderen voor rij ${rowNum}: ${e.message}`);
        Logger.log(`Could not delete calendar event for row ${rowNum}. Error: ${e.toString()}`);
      }
    } else {
      logMessage(`Max seats = 0 voor rij ${rowNum}, maar geen bestaand agenda-item gevonden om te verwijderen.`);
      Logger.log(`Max seats = 0 for row ${rowNum}, but no existing calendar event ID found to delete.`);
    }
    return; // Exit early, no need to create/update event when max seats is 0
  }
  
  if (!dateValue || !location) return;
  
  // Use a lock to prevent duplicate calendar event creation from concurrent triggers
  const lock = LockService.getScriptLock();
  const lockKey = `calendarSync_row_${rowNum}`;
  const cache = CacheService.getScriptCache();
  
  // Check if this row is already being processed (within last 10 seconds)
  const isProcessing = cache.get(lockKey);
  if (isProcessing) {
    Logger.log(`Row ${rowNum} is already being processed for calendar sync. Skipping to prevent duplicates.`);
    return;
  }
  
  // Try to acquire lock, wait max 5 seconds
  if (!lock.tryLock(5000)) {
    Logger.log(`Could not acquire lock for calendar sync on row ${rowNum}. Skipping.`);
    return;
  }
  
  try {
    // Mark this row as being processed for 10 seconds
    cache.put(lockKey, 'processing', 10);
    
    let titleSuffix = !bookedSeats ? `${location} (OPTIE - nog geen deelnemers)` : `${location} (${bookedSeats} ${bookedSeats === 1 ? 'deelnemer' : 'deelnemers'})`;
    const title = `Thermoclinic op/bij ${titleSuffix}`;
    const eventDate = new Date(dateValue);
    const calendar = CalendarApp.getCalendarById(TARGET_CALENDAR_ID);
    if (!calendar) {
      Logger.log(`FOUT: Kalender met ID ${TARGET_CALENDAR_ID} niet gevonden.`);
      logMessage(`FOUT: Kalender met ID ${TARGET_CALENDAR_ID} niet gevonden voor sync. Check configuratie.`);
      return;
    }
    let event;
    if (eventId) {
      try {
        event = calendar.getEventById(eventId);
      } catch (e) {
        Logger.log(`Could not find event with ID ${eventId}. A new one will be created. Error: ${e.message}`);
        logMessage(`WAARSCHUWING: Agenda-item met ID ${eventId} niet gevonden. Nieuw item wordt aangemaakt.`);
      }
    }
    let startTime, endTime;
    let isAllDay = false;
    if (timeValue) {
      const match = String(timeValue).match(/(\d{1,2})[:.]?(\d{2})?\s*-\s*(\d{1,2})[:.]?(\d{2})?/);
      if (match) {
        const startHour = parseInt(match[1], 10);
        const startMinute = parseInt(match[2], 10) || 0;
        let endHour = parseInt(match[3], 10);
        let endMinute = parseInt(match[4], 10) || 0;
        
        // Handle cases where start/end times might cross midnight (e.g., 22:00 - 01:00)
        startTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), startHour, startMinute, 0);
        endTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), endHour, endMinute, 0);

        // If end time is earlier than start time, assume it's on the next day
        if (endTime < startTime) {
            endTime.setDate(endTime.getDate() + 1);
        }

      } else {
        // If time format doesn't match HH:MM - HH:MM, try HH:MM
        const singleTimeMatch = String(timeValue).match(/(\d{1,2})[:.]?(\d{2})/);
        if(singleTimeMatch) {
            const hour = parseInt(singleTimeMatch[1], 10);
            const minute = parseInt(singleTimeMatch[2], 10) || 0;
            startTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), hour, minute, 0);
            endTime = new Date(startTime.getTime() + DEFAULT_EVENT_DURATION_HOURS * 60 * 60 * 1000); // Default duration
        } else {
            isAllDay = true; // Fallback to all-day if no time can be parsed
        }
      }
    } else {
      isAllDay = true;
    }
    const eventOptions = {
      location: location
    };
    if (event) {
      event.setTitle(title);
      event.setLocation(location);
      if (isAllDay) {
        event.setAllDayDate(eventDate);
      } else {
        event.setTime(startTime, endTime);
      }
      Logger.log(`Updated calendar event for "${location}" on row ${rowNum}.`);
      logMessage(`Agenda-item bijgewerkt voor "${title}" op rij ${rowNum}.`);
    } else {
      let newEvent = isAllDay ? calendar.createAllDayEvent(title, eventDate, eventOptions) : calendar.createEvent(title, startTime, endTime, eventOptions);
      sheet.getRange(rowNum, headerMap[CALENDAR_EVENT_ID_HEADER] + 1).setValue(newEvent.getId());
      Logger.log(`Created new calendar event for "${location}" on row ${rowNum}.`);
      logMessage(`Nieuw agenda-item aangemaakt voor "${title}" op rij ${rowNum}. ID: ${newEvent.getId()}`);
    }
  } finally {
    // Always release the lock
    lock.releaseLock();
  }
}

function populateCoreAppFormDropdown() {
  try {
    const form = FormApp.openById(CORE_APP_FORM_ID);
    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const sheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
    if (!sheet) return;
    let dropdownItem = form.getItems().find(item => item.getTitle() === CORE_APP_QUESTION_TITLE) ?.asListItem();
    if (!dropdownItem) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_CLINICS_START_ROW) {
      dropdownItem.setChoiceValues(['Geen actieve clinics gevonden.']);
      return;
    }
    const allData = sheet.getRange(DATA_CLINICS_START_ROW, 1, lastRow - DATA_CLINICS_START_ROW + 1, BOOKED_SEATS_COLUMN_INDEX).getValues();
    const dateTimeOptions = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    for (const rowData of allData) {
      const dateValue = rowData[DATE_COLUMN_INDEX - 1];
      if (!dateValue) continue;
      const eventDate = new Date(dateValue);
      if (eventDate < today) continue;
      const bookedSeatsRaw = rowData[BOOKED_SEATS_COLUMN_INDEX - 1];
      const numBookedSeats = (bookedSeatsRaw === '' || bookedSeatsRaw === null) ? 0 : parseInt(bookedSeatsRaw, 10);
      if (isNaN(numBookedSeats) || numBookedSeats < 1) continue;
      const timeText = String(rowData[TIME_COLUMN_INDEX - 1]).trim();
      const locationText = String(rowData[LOCATION_COLUMN_INDEX - 1]).trim();
      if (!timeText || !locationText) continue;
      const combinedOption = `${getDutchDateString(eventDate)} ${timeText}, ${locationText}`;
      dateTimeOptions.push(combinedOption);
    }
    dropdownItem.setChoiceValues(dateTimeOptions.length > 0 ? dateTimeOptions : ['Geen actieve clinics met deelnemers gevonden.']);
  } catch (err) {
    Logger.log(`populateCoreApp... ERROR: ${err.toString()}`);
    logMessage(`FOUT bij bijwerken CORE-app dropdown: ${err.message}`);
  }
}

function populateFormDropdown(formType) {
  try {
    const formId = formType.toLowerCase() === 'open' ? OPEN_FORM_ID : BESLOTEN_FORM_ID;
    const form = FormApp.openById(formId);
    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const sheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
    if (!sheet) return;
    SpreadsheetApp.flush();
    let dropdownItem = form.getItems().find(item => item.getTitle() === QUESTION_TITLE_TO_UPDATE) ?.asListItem();
    if (!dropdownItem) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_CLINICS_START_ROW) {
      dropdownItem.setChoiceValues(['Geen beschikbare datums/tijdstippen gevonden']);
      return;
    }
    const allData = sheet.getRange(DATA_CLINICS_START_ROW, 1, lastRow - DATA_CLINICS_START_ROW + 1, TYPE_COLUMN_INDEX).getValues();
    const dateTimeOptions = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    for (const rowData of allData) {
      const typeValue = rowData[TYPE_COLUMN_INDEX - 1];
      if (String(typeValue).trim().toLowerCase() !== formType.toLowerCase()) continue;
      const dateValue = rowData[DATE_COLUMN_INDEX - 1];
      if (!dateValue) continue;
      const eventDate = new Date(dateValue);
      if (eventDate <= today) continue;
      const maxSeatsRaw = rowData[MAX_SEATS_COLUMN_INDEX - 1];
      const bookedSeatsRaw = rowData[BOOKED_SEATS_COLUMN_INDEX - 1];
      const numMaxSeats = parseInt(maxSeatsRaw, 10);
      const numBookedSeats = (bookedSeatsRaw === '' || bookedSeatsRaw === null) ? 0 : parseInt(bookedSeatsRaw, 10);
      if (isNaN(numMaxSeats) || numMaxSeats <= 0 || isNaN(numBookedSeats) || numBookedSeats >= numMaxSeats) continue;
      const timeText = String(rowData[TIME_COLUMN_INDEX - 1]).trim();
      const locationText = String(rowData[LOCATION_COLUMN_INDEX - 1]).trim();
      if (!timeText || !locationText) continue;
      const availableSeats = numMaxSeats - numBookedSeats;
      const seatsText = availableSeats === 1 ? '1 plaats over' : `${availableSeats} plaatsen over`;
      const combinedOption = `${getDutchDateString(eventDate)} ${timeText}, ${locationText} (${seatsText})`;
      dateTimeOptions.push(combinedOption);
    }
    dropdownItem.setChoiceValues(dateTimeOptions.length > 0 ? dateTimeOptions : ['Momenteel zijn er geen beschikbare plaatsen.']);
  } catch (err) {
    Logger.log(`populate... ERROR for form type ${formType}: ${err.toString()}`);
    logMessage(`FOUT bij bijwerken dropdown voor formulier type ${formType}: ${err.message}`);
  }
}

/**
 * Handles participant name changes in response sheets by renaming their corresponding folders.
 * This function is triggered when someone directly edits first name or last name columns in the response sheets.
 * @param {Object} e The edit event object from the onEdit trigger.
 */
function handleParticipantNameChange(e) {
  try {
    const sheet = e.range.getSheet();
    const editedRow = e.range.getRow();
    const editedCol = e.range.getColumn();
    
    // Get headers to identify which column was edited
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const editedHeader = headers[editedCol - 1];
    
    // Only proceed if a name column was edited
    if (editedHeader !== FORM_FIRST_NAME_QUESTION_TITLE && editedHeader !== FORM_LAST_NAME_QUESTION_TITLE) {
      return; // Not a name change, exit early
    }
    
    logMessage(`Deelnemernaam gewijzigd in rij ${editedRow}, kolom '${editedHeader}'.`);
    
    // Get column indices
    const firstNameColIdx = headers.indexOf(FORM_FIRST_NAME_QUESTION_TITLE);
    const lastNameColIdx = headers.indexOf(FORM_LAST_NAME_QUESTION_TITLE);
    const participantNumberColIdx = headers.indexOf(DEELNEMERNUMMER_HEADER);
    const folderIdColIdx = headers.indexOf(DRIVE_FOLDER_ID_HEADER);
    
    // Validate required columns exist
    if (participantNumberColIdx === -1 || folderIdColIdx === -1) {
      logMessage(`FOUT: Benodigde kolommen ontbreken. ${DEELNEMERNUMMER_HEADER}: ${participantNumberColIdx}, ${DRIVE_FOLDER_ID_HEADER}: ${folderIdColIdx}`);
      return;
    }
    
    // Get participant data from the edited row
    const rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const firstName = firstNameColIdx !== -1 ? String(rowData[firstNameColIdx] || '').trim() : '';
    const lastName = lastNameColIdx !== -1 ? String(rowData[lastNameColIdx] || '').trim() : '';
    const participantNumber = String(rowData[participantNumberColIdx] || '').trim();
    const folderId = String(rowData[folderIdColIdx] || '').trim();
    
    if (!participantNumber) {
      logMessage(`FOUT: Geen deelnemernummer gevonden in rij ${editedRow}.`);
      return;
    }
    
    // Format participant number to always have 2 digits (e.g., "7" becomes "07")
    const formattedParticipantNumber = Utilities.formatString('%02d', parseInt(participantNumber) || 0);
    
    // Create new folder name
    let newFolderName;
    if (!firstName && !lastName) {
      newFolderName = `${formattedParticipantNumber} ${DEFAULT_UNASSIGNED_PARTICIPANT_NAME}`;
    } else {
      newFolderName = `${formattedParticipantNumber} ${firstName} ${lastName}`.replace(/\s+/g, ' ').trim();
    }
    
    // Validate folder name length (Google Drive limit is 255 characters)
    if (newFolderName.length > 255) {
      // Truncate last name to fit within limit
      const maxLastNameLength = 255 - participantNumber.length - firstName.length - 3; // 3 for spaces
      const truncatedLastName = lastName.substring(0, Math.max(0, maxLastNameLength));
      newFolderName = `${participantNumber} ${firstName} ${truncatedLastName}`.replace(/\s+/g, ' ').trim();
      logMessage(`WAARSCHUWING: Mapnaam ingekort tot ${newFolderName.length} karakters: "${newFolderName}"`);
    }
    
    // Handle folder operations
    if (folderId) {
      // Try to rename existing folder
      try {
        const folder = DriveApp.getFolderById(folderId);
        const currentName = folder.getName();
        
        if (currentName === newFolderName) {
          logMessage(`Info: Mapnaam is al correct: "${newFolderName}". Geen actie nodig.`);
          return;
        }
        
        // Check for duplicate names in the same parent folder
        const parentFolders = folder.getParents();
        if (parentFolders.hasNext()) {
          const parentFolder = parentFolders.next();
          const existingFolders = parentFolder.getFoldersByName(newFolderName);
          
          if (existingFolders.hasNext()) {
            // Handle duplicate by appending number
            let counter = 2;
            let uniqueName = `${newFolderName} (${counter})`;
            
            while (parentFolder.getFoldersByName(uniqueName).hasNext()) {
              counter++;
              uniqueName = `${newFolderName} (${counter})`;
              if (counter > 100) break; // Safety limit
            }
            
            newFolderName = uniqueName;
            logMessage(`WAARSCHUWING: Duplicate mapnaam gevonden. Hernoemd naar: "${newFolderName}"`);
          }
        }
        
        folder.setName(newFolderName);
        logMessage(`Deelnemersmap hernoemd van "${currentName}" naar "${newFolderName}".`);
        
      } catch (folderError) {
        // Folder no longer exists, create new one
        logMessage(`WAARSCHUWING: Bestaande map (ID: ${folderId}) niet gevonden. Nieuwe map wordt aangemaakt.`);
        
        // Find the event folder to create the new participant folder in
        const eventNameCol = headers.indexOf(FORM_EVENT_QUESTION_TITLE);
        if (eventNameCol !== -1) {
          const eventName = String(rowData[eventNameCol] || '').replace(/\s\(.*\)$/, '').trim();
          
          try {
            // Find event folder ID from Data Clinics sheet
            const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
            const dataClinicsSheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
            const dataClinicsData = dataClinicsSheet.getDataRange().getValues();
            const dataClinicsHeaders = dataClinicsData.shift();
            const eventFolderIdColIdx = dataClinicsHeaders.indexOf(EVENT_FOLDER_ID_HEADER);
            
            let eventFolderId = null;
            for (const row of dataClinicsData) {
              const dateValue = row[DATE_COLUMN_INDEX - 1];
              if (!dateValue) continue;
              
              const reconstructedName = `${getDutchDateString(new Date(dateValue))} ${String(row[TIME_COLUMN_INDEX - 1] || '').trim()}, ${String(row[LOCATION_COLUMN_INDEX - 1] || '').trim()}`;
              if (reconstructedName === eventName && eventFolderIdColIdx !== -1) {
                eventFolderId = row[eventFolderIdColIdx];
                break;
              }
            }
            
            if (eventFolderId) {
              const eventFolder = DriveApp.getFolderById(eventFolderId);
              
              // Check for duplicate names and handle
              const existingFolders = eventFolder.getFoldersByName(newFolderName);
              if (existingFolders.hasNext()) {
                let counter = 2;
                let uniqueName = `${newFolderName} (${counter})`;
                
                while (eventFolder.getFoldersByName(uniqueName).hasNext()) {
                  counter++;
                  uniqueName = `${newFolderName} (${counter})`;
                  if (counter > 100) break;
                }
                
                newFolderName = uniqueName;
              }
              
              const newFolder = eventFolder.createFolder(newFolderName);
              const newFolderId = newFolder.getId();
              
              // Update the Participant Folder ID in the sheet
              sheet.getRange(editedRow, folderIdColIdx + 1).setValue(newFolderId);
              
              logMessage(`Nieuwe deelnemersmap aangemaakt: "${newFolderName}" (ID: ${newFolderId})`);
              
            } else {
              logMessage(`FOUT: Kon event folder niet vinden voor "${eventName}".`);
            }
            
          } catch (createError) {
            logMessage(`FOUT bij aanmaken nieuwe map: ${createError.message}`);
            Logger.log(`Create folder error: ${createError.toString()}`);
          }
        }
      }
      
    } else {
      logMessage(`Info: Geen Participant Folder ID gevonden voor rij ${editedRow}. Geen mapactie ondernomen.`);
    }
    
  } catch (err) {
    logMessage(`FOUT in handleParticipantNameChange: ${err.message}`);
    Logger.log(`handleParticipantNameChange ERROR: ${err.toString()}\n${err.stack}`);
  } finally {
    flushLogs();
  }
}