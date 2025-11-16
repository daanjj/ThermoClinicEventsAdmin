// --- START OF FILE FormSubmission v2.4.js (Corrected) ---

// This file contains the logic for processing new form submissions (bookings)
// from the open and besloten forms.

function processBooking(e) {
  const fromAlias = "info@thermoclinics.nl";
  try {
    if (!e || !e.namedValues) {
        logMessage(`processBooking ERROR: No event object or namedValues received.`);
        return;
    }

    const selectedEventOptionWithSeats = e.namedValues[FORM_EVENT_QUESTION_TITLE][0];
    const selectedEventOption = selectedEventOptionWithSeats.replace(/\s\(.*\)$/, '').trim();

    const eventParts = selectedEventOption.split(',');
    const location = eventParts.length > 1 ? eventParts[1].trim() : '';
    const dateTimePart = eventParts[0].trim();
    const timeRegexMatch = dateTimePart.match(/(\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2})|(\d{1,2}:\d{2})/);
    const time = timeRegexMatch ? timeRegexMatch[0].trim() : ''; 
    const date = timeRegexMatch ? dateTimePart.substring(0, timeRegexMatch.index).trim() : dateTimePart; 

    const placeholderMap = {
      '<Voornaam>': String(e.namedValues[FORM_FIRST_NAME_QUESTION_TITLE][0] || '').trim(),
      '<Achternaam>': String(e.namedValues[FORM_LAST_NAME_QUESTION_TITLE][0] || '').trim(),
      '<Email>': String(e.namedValues[FORM_EMAIL_QUESTION_TITLE][0] || '').trim(),
      '<Eventnaam>': selectedEventOption,
      '<Locatie>': location,
      '<Datum>': date,
      '<Tijd>': time,
      '<Telefoonnummer>': String(e.namedValues[FORM_PHONE_QUESTION_TITLE]?.[0] || '').trim(),
      '<Geboortedatum>': String(e.namedValues[FORM_DOB_QUESTION_TITLE]?.[0] || '').trim(),
      '<Woonplaats>': String(e.namedValues[FORM_CITY_QUESTION_TITLE]?.[0] || '').trim()
    };
    
    // ... (Your existing logic for finding the row, updating seats, and creating folders remains the same)
    let foundRowIndex = -1;
    let participantSequenceNumber = 'XX';
    let participantSubfolderId = '';

    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const dataClinicsSheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
    if (!dataClinicsSheet) throw new Error(`Data Clinics sheet not found.`);
    SpreadsheetApp.flush(); 

    const allData = dataClinicsSheet.getDataRange().getValues();
    const headers = allData.shift();
    const eventFolderIdColIdx = headers.indexOf(EVENT_FOLDER_ID_HEADER);

    let clinicType = '';
    for (let i = 0; i < allData.length; i++) {
      const rowData = allData[i];
      const sheetDateValue = rowData[DATE_COLUMN_INDEX - 1];
      if (!sheetDateValue) continue;
      
      const reconstructedOption = `${getDutchDateString(new Date(sheetDateValue))} ${String(rowData[TIME_COLUMN_INDEX - 1]).trim()}, ${String(rowData[LOCATION_COLUMN_INDEX - 1]).trim()}`;
      
      if (reconstructedOption === selectedEventOption) {
        foundRowIndex = i + DATA_CLINICS_START_ROW;
        clinicType = String(rowData[TYPE_COLUMN_INDEX - 1] || '').trim().toLowerCase();
        const currentBookedSeats = (rowData[BOOKED_SEATS_COLUMN_INDEX - 1] || 0);
        const newBookedSeats = currentBookedSeats + 1;
        dataClinicsSheet.getRange(foundRowIndex, BOOKED_SEATS_COLUMN_INDEX).setValue(newBookedSeats);
        participantSequenceNumber = Utilities.formatString('%02d', newBookedSeats);
        placeholderMap['<Deelnemernummer>'] = participantSequenceNumber;

        const eventDateFormatted = Utilities.formatDate(new Date(sheetDateValue), FORMATTING_TIME_ZONE, DATE_FORMAT_YYYYMMDD);
        const timeFormatted = String(rowData[TIME_COLUMN_INDEX - 1]).trim().replace(/:|\./g, '');
        const eventFolderName = `${eventDateFormatted} ${timeFormatted} ${String(rowData[LOCATION_COLUMN_INDEX-1]).trim()}`;
        
        const parentFolder = DriveApp.getFolderById(PARENT_EVENT_FOLDER_ID);
        const folders = parentFolder.getFoldersByName(eventFolderName);
        const eventFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(eventFolderName);

        if (eventFolderIdColIdx !== -1 && !rowData[eventFolderIdColIdx]) {
          dataClinicsSheet.getRange(foundRowIndex, eventFolderIdColIdx + 1).setValue(eventFolder.getId());
        }

        const participantSubfolderName = `${participantSequenceNumber} ${placeholderMap['<Voornaam>']} ${placeholderMap['<Achternaam>']}`.trim();
        const subfolder = eventFolder.createFolder(participantSubfolderName);
        participantSubfolderId = subfolder.getId();
        break;
      }
    }
    
    if (foundRowIndex === -1) throw new Error(`Event "${selectedEventOption}" not found.`);
    
    syncCalendarEventFromSheet(foundRowIndex);

    const targetSheet = e.range.getSheet();
    const targetRow = e.range.getRow();
    const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    const pNumCol = targetHeaders.indexOf(DEELNEMERNUMMER_HEADER);
    const fIdCol = targetHeaders.indexOf(DRIVE_FOLDER_ID_HEADER);
    
    if (pNumCol !== -1) targetSheet.getRange(targetRow, pNumCol + 1).setValue(participantSequenceNumber);
    if (fIdCol !== -1) targetSheet.getRange(targetRow, fIdCol + 1).setValue(participantSubfolderId);

    // Send confirmation email with appropriate template based on clinic type
    if (placeholderMap['<Email>']) {
      // Determine which template to use based on clinic type
      let templateId;
      let templateType;
      
      if (clinicType === 'open') {
        templateId = OPEN_CONFIRMATION_EMAIL_TEMPLATE_ID;
        templateType = 'Open';
      } else if (clinicType === 'besloten') {
        templateId = BESLOTEN_CONFIRMATION_EMAIL_TEMPLATE_ID;
        templateType = 'Besloten';
      } else {
        // Fallback to original template if type is not recognized
        templateId = CONFIRMATION_EMAIL_TEMPLATE_ID;
        templateType = 'Default';
        logMessage(`WAARSCHUWING: Onbekend clinic type '${clinicType}'. Gebruik standaard template.`);
      }
      
      const mergedMail = mergeSingleTemplate(templateId, placeholderMap);
      
      GmailApp.sendEmail(placeholderMap['<Email>'], mergedMail.subject, '', { name: mergedMail.senderName, htmlBody: mergedMail.htmlBody, from: fromAlias });
      
      const templateName = DriveApp.getFileById(templateId).getName();
      logMessage(`${templateType} registratiebevestiging verstuurd aan: ${placeholderMap['<Email>']}, Onderwerp: "${mergedMail.subject}"`);
    } else {
        logMessage(`WAARSCHUWING: Geen e-mailadres. Bevestigingsmail niet verstuurd.`);
    }
    
    updateAllFormDropdowns();

  } catch (err) {
    Logger.log(`processBooking CRITICAL ERROR: ${err.toString()}\n${err.stack}`);
    logMessage(`processBooking CRITICAL ERROR: ${err.message}`);
  } finally {
    flushLogs();
  }
}