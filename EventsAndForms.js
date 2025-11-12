// This file handles logic related to clinic events (time changes, calendar sync)
// and populating form dropdowns based on clinic data.

/**
 * Triggered on an edit in 'Data clinics'. If the 'Tijdstip' column is changed,
 * this function updates corresponding entries in response sheets AND renames the event folder in Drive.
 * @param {Object} e The event object from the onEdit trigger.
 */
function handleTimeChange(e) {
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
  
  // Check if the edited column is 'Tijdstip'
  if (e.range.getColumn() !== TIME_COLUMN_INDEX) {
    return;
  }
  
  logMessage(`Tijdstip wijziging gedetecteerd in '${DATA_CLINICS_SHEET_NAME}' op rij ${editedRow}. Oud: "${e.oldValue}", Nieuw: "${e.value}".`);

  // --- Step 2: Read row data and construct event names ---
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const eventFolderIdColIdx = headers.indexOf(EVENT_FOLDER_ID_HEADER);
  
  const rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateValue = rowData[DATE_COLUMN_INDEX - 1];
  const locationValue = rowData[LOCATION_COLUMN_INDEX - 1];
  const eventFolderId = (eventFolderIdColIdx !== -1) ? rowData[eventFolderIdColIdx] : null;

  if (!dateValue || !locationValue) {
    logMessage(`FOUT: Kan event niet identificeren. Datum of locatie ontbreekt op rij ${editedRow}.`);
    return;
  }

  // --- Step 3: Rename the Google Drive Folder ---
  if (eventFolderId) {
    try {
      const newTimeFormatted = String(e.value).trim().replace(/:|\./g, '');
      const newFolderName = `${Utilities.formatDate(new Date(dateValue), FORMATTING_TIME_ZONE, DATE_FORMAT_YYYYMMDD)} ${newTimeFormatted} ${locationValue}`;
      const folder = DriveApp.getFolderById(eventFolderId);
      folder.setName(newFolderName);
      logMessage(`Event folder hernoemd naar: "${newFolderName}".`);
    } catch (driveError) {
      logMessage(`WAARSCHUWING: Kon de event folder met ID "${eventFolderId}" niet hernoemen. Fout: ${driveError.message}`);
      Logger.log(`Could not rename folder with ID ${eventFolderId}. Error: ${driveError.toString()}`);
    }
  } else {
    logMessage(`Info: Geen Event Folder ID gevonden op rij ${editedRow}. Hernoemen overgeslagen.`);
  }

  // --- Step 4: Update participant entries in response sheets ---
  const dateString = getDutchDateString(new Date(dateValue));
  const oldEventNameBase = `${dateString} ${String(e.oldValue).trim()}, ${String(locationValue).trim()}`;
  const newEventNameBase = `${dateString} ${String(e.value).trim()}, ${String(locationValue).trim()}`;
  
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
  const dateValue = rowData[headerMap['Datum']];
  const timeValue = rowData[headerMap['Tijdstip']];
  const location = rowData[headerMap['Locatie']];
  const bookedSeatsRaw = rowData[headerMap['Aantal boekingen']];
  const eventId = rowData[headerMap[CALENDAR_EVENT_ID_HEADER]];

  const bookedSeats = parseInt(bookedSeatsRaw, 10) || 0;
  if (!dateValue || !location) return;
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