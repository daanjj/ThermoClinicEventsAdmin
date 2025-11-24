// This file contains general utility functions used across the script.

let logBuffer = []; // Global buffer for efficient logging

/**
 * UTILITY FUNCTION: This function scans the 'Data clinics' sheet for events
 * that are missing an 'Event Folder ID'. For each missing one, it attempts to
 * find the corresponding folder in Drive and back-fills the ID.
 * To run this, select it from the function list in the Apps Script editor and click "Run".
 */
function updateEventFolderIDs() {
  const feedback = [];
  let updatedCount = 0;
  let notFoundCount = 0;

  try {
    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const sheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet '${DATA_CLINICS_SHEET_NAME}' not found.`);
    }

    const allData = sheet.getDataRange().getValues();
    const headers = allData.shift();
    
    const dateColIdx = headers.indexOf('Datum');
    const timeColIdx = headers.indexOf('Tijdstip');
    const locationColIdx = headers.indexOf('Locatie');
    const eventFolderIdColIdx = headers.indexOf(EVENT_FOLDER_ID_HEADER);

    if (eventFolderIdColIdx === -1) {
      throw new Error(`The required column '${EVENT_FOLDER_ID_HEADER}' was not found in the sheet.`);
    }

    const parentFolder = DriveApp.getFolderById(PARENT_EVENT_FOLDER_ID);
    feedback.push(`Scanning ${allData.length} events in '${DATA_CLINICS_SHEET_NAME}'...`);

    allData.forEach((row, index) => {
      const currentFolderId = row[eventFolderIdColIdx];
      const rowNum = index + DATA_CLINICS_START_ROW;
      
      if (!currentFolderId) {
        const dateValue = row[dateColIdx];
        const timeValue = row[timeColIdx];
        const locationValue = row[locationColIdx];

        if (!dateValue || !timeValue || !locationValue) {
          feedback.push(`- Rij ${rowNum}: Overgeslagen, onvoldoende data om mapnaam te construeren.`);
          return;
        }

        const dateForFolderName = Utilities.formatDate(new Date(dateValue), FORMATTING_TIME_ZONE, DATE_FORMAT_YYYYMMDD);
        const timeForFolderName = String(timeValue).trim().replace(/:|\./g, '');
        const folderName = `${dateForFolderName} ${timeForFolderName} ${locationValue}`;
        
        const folders = parentFolder.getFoldersByName(folderName);
        if (folders.hasNext()) {
          const foundFolder = folders.next();
          const foundFolderId = foundFolder.getId();
          sheet.getRange(rowNum, eventFolderIdColIdx + 1).setValue(foundFolderId);
          updatedCount++;
          feedback.push(`+ Rij ${rowNum}: Map gevonden voor "${folderName}". ID ${foundFolderId} ingevuld.`);
        } else {
          notFoundCount++;
          feedback.push(`- Rij ${rowNum}: Map NIET GEVONDEN met de naam "${folderName}".`);
        }
      }
    });

    let summaryMessage = `Scannen voltooid.\n\n${updatedCount} Event Folder ID's gevonden en bijgewerkt.`;
    if (notFoundCount > 0) {
      summaryMessage += `\n${notFoundCount} event folders konden niet worden gevonden in Drive.`;
    }
    summaryMessage += `\n\nBekijk het logboek voor gedetailleerde resultaten.`;
    
    SpreadsheetApp.getUi().alert(summaryMessage);

  } catch (e) {
    Logger.log(`ERROR in updateEventFolderIDs: ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`Er is een fout opgetreden: ${e.message}`);
    feedback.push(`FATALE FOUT: ${e.message}`);
  } finally {
    logMessage('----- START Utility: updateEventFolderIDs -----');
    feedback.forEach(line => logMessage(line));
    logMessage('----- EINDE Utility: updateEventFolderIDs -----');
    flushLogs();
  }
}

/**
 * This function is designed to be run by a user from the menu.
 * Its primary purpose is to trigger the Google Apps Script authorization flow,
 * requesting all necessary permissions at once. It also provides feedback to the user.
 * This version is more robust and provides better error logging.
 */
function forceAuthorization() {
  try {
    // --- Standard Service Checks ---
    SpreadsheetApp.getActiveSpreadsheet().getName();
    DriveApp.getRootFolder().getName();
    GmailApp.getAliases();
    UrlFetchApp.fetch('https://www.google.com', {
      muteHttpExceptions: true
    });

    // --- User Info Check (for Session.getEffectiveUser) ---
    try {
      const userEmail = Session.getEffectiveUser().getEmail();
      Logger.log(`User info check passed. Active user: ${userEmail}`);
    } catch (userInfoError) {
      throw new Error("De userinfo.email permissie is niet toegekend. Dit is nodig voor account verificatie bij mail merge. Fout: " + userInfoError.message);
    }

    // --- Advanced Drive Service Check ---
    try {
      Drive.About.get();
    } catch (driveError) {
      throw new Error("De geavanceerde Google Drive-service is mislukt. Controleer of de 'Google Drive API' is ingeschakeld in de Google Cloud Console. Fout: " + driveError.message);
    }

    // --- Calendar Service Check (with write operation) ---
    // This block now performs a write action (create/delete event) to force the correct
    // authorization scope (https://www.googleapis.com/auth/calendar).
    let tempEvent = null;
    try {
      const calendar = CalendarApp.getDefaultCalendar();
      const eventTitle = 'Temporary Authorization Check - This will be deleted instantly';
      const startTime = new Date();
      const endTime = new Date(startTime.getTime() + (1 * 60 * 1000)); // 1 minute duration

      tempEvent = calendar.createEvent(eventTitle, startTime, endTime);
      // The creation itself is what triggers the auth prompt.
      // If it succeeds, we've done our job.

    } catch (calendarError) {
      // This error is expected to happen during the first run when the auth prompt appears.
      // The user grants permission, the script times out, and they need to run it again.
      throw new Error("De Google Agenda-service is mislukt. Dit is normaal tijdens de eerste autorisatie. Geef toestemming en probeer de functie daarna opnieuw. Fout: " + calendarError.message);
    } finally {
      // CRUCIAL: Always delete the temporary event, even if an error occurred.
      if (tempEvent) {
        try {
          tempEvent.deleteEvent();
        } catch (deleteError) {
          // If deletion fails, log it but don't bother the user.
          Logger.log("Could not delete temporary calendar event. Please delete it manually. Title: " + tempEvent.getTitle());
        }
      }
    }

    // If the script reaches this point, everything is working.
    SpreadsheetApp.getUi().alert('Controle voltooid. Alle benodigde permissies en services zijn actief. U kunt nu de andere functies gebruiken.');

  } catch (e) {
    // This will now show a much more specific error message to the user.
    SpreadsheetApp.getUi().alert('Er is een fout opgetreden: ' + e.message);
  }
}

/**
 * Escapes special characters in a string for use in a regular expression.
 * @param {string} str The string to escape.
 * @returns {string} The escaped string.
 */
function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Finds and resolves time arithmetic placeholders like "<Tijd + 30 min>" in a string.
 * @param {string} text The email body text containing placeholders.
 * @param {Object} placeholderMap A map of existing placeholder values (e.g., { '<Tijd>': '12:00' }).
 * @returns {string} The text with time arithmetic placeholders replaced by calculated times.
 */
function resolveTimeArithmeticPlaceholders(text, placeholderMap) {
  const timeRegex = /(?:<|<)([A-Za-z_]+)\s*([+-])\s*(\d+)\s*min(?:>|>)/g;

  return text.replace(timeRegex, (match, basePlaceholderKey, operator, minutesStr) => {
    const fullBasePlaceholder = `<${basePlaceholderKey}>`;
    const baseTimeValue = placeholderMap[fullBasePlaceholder];

    if (!baseTimeValue) {
      Logger.log(`Time Arithmetic: Base placeholder ${fullBasePlaceholder} not found. Skipping calculation.`);
      return match;
    }

    const timeParts = baseTimeValue.split(/[:.]/);
    const hours = parseInt(timeParts[0], 10);
    const minutes = parseInt(timeParts[1], 10) || 0;

    if (isNaN(hours) || isNaN(minutes)) {
      Logger.log(`Time Arithmetic: Could not parse base time "${baseTimeValue}". Skipping calculation.`);
      return match;
    }

    try {
      const date = new Date();
      date.setHours(hours, minutes, 0, 0);
      const offset = parseInt(minutesStr, 10);

      if (operator === '+') {
        date.setMinutes(date.getMinutes() + offset);
      } else {
        date.setMinutes(date.getMinutes() - offset);
      }

      const newHours = date.getHours().toString().padStart(2, '0');
      const newMinutes = date.getMinutes().toString().padStart(2, '0');

      return `${newHours}:${newMinutes}`;
    } catch (e) {
      Logger.log(`Time Arithmetic: Error during calculation for "${baseTimeValue}". Error: ${e.toString()}`);
      return match;
    }
  });
}

/**
 * Adds a log message to a global buffer instead of writing it directly.
 * This is highly efficient as it avoids API calls in loops.
 * @param {string} message The message to log.
 */
function logMessage(message) {
  const timestamp = Utilities.formatDate(new Date(), FORMATTING_TIME_ZONE, "yyyy-MM-dd HH:mm:ss");
  logBuffer.push(`${timestamp} - ${message}`);
}

/**
 * Writes all buffered log messages to the Google Doc in a single operation.
 * This should be called only once, at the end of a script's execution.
 */
function flushLogs() {
  if (logBuffer.length === 0) {
    return; // Nothing to log
  }

  try {
    // Join all messages with a newline character for a single write
    const fullLogText = logBuffer.join('\n');

    const doc = DocumentApp.openById(LOG_DOCUMENT_ID);
    const body = doc.getBody();
    body.appendParagraph(fullLogText);

    // Clear the buffer for the next script run
    logBuffer = [];
  } catch (e) {
    // If logging to the doc fails, log to the built-in logger as a fallback
    Logger.log(`Failed to write to log document ID ${LOG_DOCUMENT_ID}: ${e.toString()}`);
    Logger.log(`Buffered Logs: \n${logBuffer.join('\n')}`);
  }
}

function logToDocument(message) {
  try {
    const doc = DocumentApp.openById(LOG_DOCUMENT_ID);
    const body = doc.getBody();
    const timestamp = Utilities.formatDate(new Date(), FORMATTING_TIME_ZONE, "yyyy-MM-dd HH:mm:ss");
    body.appendParagraph(`${timestamp} - ${message}`);
  } catch (e) {
    Logger.log(`Failed to write to log document ID ${LOG_DOCUMENT_ID}: ${e.toString()}`);
  }
}

function getDutchDateString(dateObject) {
  if (!dateObject || !(dateObject instanceof Date) || isNaN(dateObject.getTime())) {
    return 'onbekende datum';
  }
  const dayNamesDutch = ["zondag", "maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag"];
  const monthNamesDutch = ["januari", "februari", "maart", "april", "mei", "juni", "juli", "augustus", "september", "oktober", "november", "december"];

  const dayName = dayNamesDutch[dateObject.getDay()];
  const dayOfMonth = dateObject.getDate();
  const monthName = monthNamesDutch[dateObject.getMonth()];
  const year = dateObject.getFullYear();

  return `${dayName} ${dayOfMonth} ${monthName} ${year}`;
}