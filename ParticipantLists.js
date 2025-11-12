// This file provides functions for generating and displaying participant lists.

/**
 * Shows a dialog to the user to select a clinic for which to generate a participant list.
 */
function showParticipantListDialog() {
  const htmlTemplate = HtmlService.createTemplateFromFile('ParticipantListDialog');
  htmlTemplate.clinics = getClinicsForParticipantList();
  SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setWidth(600).setHeight(500), 'Maak deelnemerslijst');
}

/**
 * Retrieves a list of all clinics that are scheduled from 30 days ago onwards
 * and have one or more participants.
 * @returns {string[]} An array of clinic names.
 */
function getClinicsForParticipantList() {
  try {
    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const sheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
    if (!sheet) {
        logMessage(`FOUT: Sheet '${DATA_CLINICS_SHEET_NAME}' niet gevonden voor deelnemerslijst.`);
        return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_CLINICS_START_ROW) return [];

    // Read only relevant columns for performance
    const maxColumnIndexToRead = Math.max(DATE_COLUMN_INDEX, TIME_COLUMN_INDEX, LOCATION_COLUMN_INDEX, BOOKED_SEATS_COLUMN_INDEX);
    const allData = sheet.getRange(DATA_CLINICS_START_ROW, 1, lastRow - DATA_CLINICS_START_ROW + 1, maxColumnIndexToRead).getValues();
    const clinicOptions = [];

    // --- CHANGE STARTS HERE ---
    // Calculate the date 30 days ago and normalize it to the start of the day.
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    thirtyDaysAgo.setHours(0, 0, 0, 0); 
    // --- CHANGE ENDS HERE ---

    for (let i = 0; i < allData.length; i++) {
      const rowData = allData[i];
      const dateValue = rowData[DATE_COLUMN_INDEX - 1];
      const bookedSeatsRaw = rowData[BOOKED_SEATS_COLUMN_INDEX - 1];

      if (!dateValue || bookedSeatsRaw === '' || bookedSeatsRaw === null) continue; // Skip if date or booked seats are missing

      let eventDate = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
      
      // --- CHANGE STARTS HERE ---
      // Compare the event date with the 'thirtyDaysAgo' date instead of 'today'.
      if (isNaN(eventDate.getTime()) || eventDate < thirtyDaysAgo) continue; // Skip invalid or dates older than 30 days
      // --- CHANGE ENDS HERE ---

      const numBookedSeats = parseInt(bookedSeatsRaw, 10);
      if (isNaN(numBookedSeats) || numBookedSeats < 1) continue; // Only clinics with participants

      const timeText = rowData[TIME_COLUMN_INDEX - 1];
      const locationText = rowData[LOCATION_COLUMN_INDEX - 1];
      const combinedOption = `${getDutchDateString(eventDate)} ${String(timeText || '').trim()}, ${String(locationText || '').trim()}`;
      clinicOptions.push(combinedOption);
    }
    return clinicOptions;
  } catch (err) {
    Logger.log(`getClinicsForParticipantList ERROR: ${err.toString()}`);
    logMessage(`FOUT bij ophalen clinics voor deelnemerslijst: ${err.message}`);
    return [];
  }
}

/**
 * Generates an HTML table of participants for a selected clinic.
 * The table shows the participant's folder name, their CORE-mailadres, and their regular email address.
 * The link is made by matching the subfolder ID in Drive with the 'Participant Folder ID' in the response sheet.
 * @param {string} selectedClinic The full name of the clinic selected by the user.
 * @returns {string} An HTML string representing the participant table or an error message.
 */
function generateParticipantTable(selectedClinic) {
  try {
    // --- STEP 1: Find clinic details (type and event folder ID) ---
    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const dataClinicsSheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
    const allDataClinicsData = dataClinicsSheet.getDataRange().getValues();
    const headers = allDataClinicsData.shift();

    let clinicType = '';
    let eventFolderId = '';
    const eventFolderIdColIdx = headers.indexOf(EVENT_FOLDER_ID_HEADER);

    if (eventFolderIdColIdx === -1) {
        throw new Error(`De benodigde kolom '${EVENT_FOLDER_ID_HEADER}' ontbreekt in het tabblad '${DATA_CLINICS_SHEET_NAME}'.`);
    }

    for (const row of allDataClinicsData) {
      const dateValue = row[DATE_COLUMN_INDEX - 1];
      if (!dateValue) continue;

      const reconstructedName = `${getDutchDateString(new Date(dateValue))} ${String(row[TIME_COLUMN_INDEX - 1] || '').trim()}, ${String(row[LOCATION_COLUMN_INDEX - 1] || '').trim()}`;
      if (reconstructedName === selectedClinic) {
        clinicType = String(row[TYPE_COLUMN_INDEX - 1] || '').trim().toLowerCase();
        eventFolderId = row[eventFolderIdColIdx];
        break;
      }
    }

    if (!clinicType) {
      throw new Error(`Kon de details (type) voor clinic "${selectedClinic}" niet vinden in het '${DATA_CLINICS_SHEET_NAME}' tabblad.`);
    }
    if (!eventFolderId) {
      throw new Error(`De '${EVENT_FOLDER_ID_HEADER}' is niet ingevuld voor clinic "${selectedClinic}" in het '${DATA_CLINICS_SHEET_NAME}' tabblad. Zonder deze ID kan de deelnemersmap niet worden gevonden.`);
    }

    // --- STEP 2: Pre-load participant data from the correct response sheet, keyed by Folder ID ---
    const participantDataMap = {}; // Key: Participant Folder ID, Value: { coreEmail: '...', regularEmail: '...' }
    const responseSheetName = clinicType === 'open' ? OPEN_FORM_RESPONSE_SHEET_NAME : BESLOTEN_FORM_RESPONSE_SHEET_NAME;
    const responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(responseSheetName);
    if (!responseSheet) {
      throw new Error(`Het respons-tabblad '${responseSheetName}' kon niet worden gevonden.`);
    }

    const responseData = responseSheet.getDataRange().getValues();
    if (responseData.length < 2) {
        return `<p>Geen responsen gevonden in het tabblad '${responseSheetName}' voor deze clinic.</p>`;
    }
    const responseHeaders = responseData.shift();
    const eventColIdx = responseHeaders.indexOf(FORM_EVENT_QUESTION_TITLE);
    const folderIdColIdx = responseHeaders.indexOf(DRIVE_FOLDER_ID_HEADER);
    const coreMailColIdx = responseHeaders.indexOf(FORM_CORE_MAIL_HEADER);
    const emailColIdx = responseHeaders.indexOf(FORM_EMAIL_QUESTION_TITLE);
    const firstNameColIdx = responseHeaders.indexOf(FORM_FIRST_NAME_QUESTION_TITLE);
    const lastNameColIdx = responseHeaders.indexOf(FORM_LAST_NAME_QUESTION_TITLE);

    const requiredCols = {
        [FORM_EVENT_QUESTION_TITLE]: eventColIdx,
        [DRIVE_FOLDER_ID_HEADER]: folderIdColIdx,
        [FORM_CORE_MAIL_HEADER]: coreMailColIdx,
        [FORM_EMAIL_QUESTION_TITLE]: emailColIdx,
        [FORM_FIRST_NAME_QUESTION_TITLE]: firstNameColIdx,
        [FORM_LAST_NAME_QUESTION_TITLE]: lastNameColIdx
    };

    for(const colName in requiredCols) {
        if (requiredCols[colName] === -1) {
            throw new Error(`De benodigde kolom '${colName}' ontbreekt in het tabblad '${responseSheetName}'.`);
        }
    }
    
    responseData.forEach(row => {
      const eventNameInSheet = (row[eventColIdx] || '').replace(/\s\(.*\)$/, '').trim();
      if (eventNameInSheet === selectedClinic) {
        const participantFolderId = String(row[folderIdColIdx] || '').trim();
        const coreEmail = String(row[coreMailColIdx] || 'N.v.t.').trim(); // Default for missing CORE-mail
        const regularEmail = String(row[emailColIdx] || 'Niet ingevuld').trim();
        const firstName = String(row[firstNameColIdx] || '').trim();
        const lastName = String(row[lastNameColIdx] || '').trim();
        
        if (participantFolderId) {
          participantDataMap[participantFolderId] = { coreEmail, regularEmail, firstName, lastName };
        }
      }
    });

    // --- STEP 3: Get participant folders from Drive and build the table data ---
    const tableData = [];
    const eventFolder = DriveApp.getFolderById(eventFolderId);
    const subfolders = eventFolder.getFolders();
    const participantFolderRegex = /^(\d+)\s.*/; // Matches folder names starting with one or more digits

    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      const folderName = subfolder.getName();
      const match = folderName.match(participantFolderRegex); // Check if it's a numbered participant folder

      if (match) {
        const subfolderId = subfolder.getId();
        const participantInfo = participantDataMap[subfolderId];
        
        tableData.push({
          folderName: folderName,
          coreEmail: participantInfo ? participantInfo.coreEmail : 'Niet gevonden in sheet',
          regularEmail: participantInfo ? participantInfo.regularEmail : 'Niet gevonden in sheet',
          sortKey: participantInfo ? `${participantInfo.firstName || ''} ${participantInfo.lastName || ''}`.trim() : folderName
        });
      }
    }
    
    // Sort by folder name (which includes participant number)
    tableData.sort((a, b) => a.folderName.localeCompare(b.folderName, undefined, { numeric: true, sensitivity: 'base' }));

    // --- STEP 4: Generate the final HTML table with three columns ---
    if (tableData.length === 0) {
      return `<p>Geen deelnemersmappen gevonden in de Drive-map voor dit event die met een nummer beginnen, of geen overeenkomende deelnemers in het respons-tabblad.</p>`;
    }

    let html = `<h3>Deelnemerslijst voor "${selectedClinic}"</h3>`;
    html += '<style>table { width: 100%; border-collapse: collapse; } th, td { border: 1px solid #ddd; padding: 8px; text-align: left; } th { background-color: #f2f2f2; }</style>';
    html += '<table>';
    html += '<tr><th>Deelnemer</th><th>CORE-mailadres</th><th>Email</th></tr>';
    tableData.forEach(item => {
      html += `<tr><td>${item.folderName}</td><td>${item.coreEmail}</td><td>${item.regularEmail}</td></tr>`;
    });
    html += '</table>';
    
    return html;

  } catch (e) {
    Logger.log(`generateParticipantTable ERROR: ${e.toString()}\n${e.stack}`);
    logMessage(`FOUT bij het genereren van de deelnemerslijst voor "${selectedClinic}": ${e.message}`);
    // Provide a user-friendly error message in the dialog
    return `<p style="color: red;">Fout bij het genereren van de lijst: ${e.message}</p><p>Controleer de logs voor meer details.</p>`;
  } finally {
    flushLogs();
  }
}