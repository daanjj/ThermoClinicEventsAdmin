// --- START OF FILE FormSubmission v2.4.js (Corrected) ---

// This file contains the logic for processing new form submissions (bookings)
// from the open and besloten forms.

function processBooking(e) {
  // Verify Gmail alias is available
  const desiredAlias = "info@thermoclinics.nl";
  const availableAliases = GmailApp.getAliases();
  const fromAlias = availableAliases.includes(desiredAlias) ? desiredAlias : null;
  
  if (!fromAlias) {
    logMessage(`WAARSCHUWING: Alias "${desiredAlias}" niet gevonden voor bevestigingsmail. Beschikbare aliassen: ${availableAliases.join(', ')}. Email wordt verstuurd zonder 'from' alias.`);
  }
  try {
    if (!e || !e.namedValues) {
        logMessage(`processBooking ERROR: No event object or namedValues received.`);
        return;
    }

    // Debug: log all available form fields
    logMessage(`Form submission received with fields: ${Object.keys(e.namedValues).join(', ')}`);

    const selectedEventOptionWithSeats = e.namedValues[FORM_EVENT_QUESTION_TITLE]?.[0];
    if (!selectedEventOptionWithSeats) {
        logMessage(`processBooking ERROR: No event selected in form submission. Form question title: "${FORM_EVENT_QUESTION_TITLE}"`);
        logMessage(`Available form fields: ${JSON.stringify(Object.keys(e.namedValues))}`);
        logMessage(`Full namedValues: ${JSON.stringify(e.namedValues)}`);
        return;
    }
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
      '<Woonplaats>': String(e.namedValues[FORM_CITY_QUESTION_TITLE]?.[0] || '').trim(),
      '<Opmerking>': String(e.namedValues[FORM_OPMERKING_QUESTION_TITLE]?.[0] || '').trim(),
      '<Motivatie>': String(e.namedValues[FORM_MOTIVATIE_QUESTION_TITLE]?.[0] || '').trim()
    };
    
    // ... (Your existing logic for finding the row, updating seats, and creating folders remains the same)
    let foundRowIndex = -1;
    let participantSequenceNumber = 'XX';
    let participantSubfolderId = '';
    let isDuplicate = false; // Track if this is a duplicate submission

    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const dataClinicsSheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
    if (!dataClinicsSheet) throw new Error(`Data Clinics sheet not found.`);
    SpreadsheetApp.flush(); 

    const allData = dataClinicsSheet.getDataRange().getValues();
    const headers = allData.shift();
    const eventFolderIdColIdx = headers.indexOf(EVENT_FOLDER_ID_HEADER);

    let clinicType = '';
    
    // First, determine the clinic type for this event by looking it up in Data Clinics
    for (let i = 0; i < allData.length; i++) {
      const rowData = allData[i];
      const sheetDateValue = rowData[DATE_COLUMN_INDEX - 1];
      if (!sheetDateValue) continue;
      
      const reconstructedOption = `${getDutchDateString(new Date(sheetDateValue))} ${String(rowData[TIME_COLUMN_INDEX - 1]).trim()}, ${String(rowData[LOCATION_COLUMN_INDEX - 1]).trim()}`;
      
      if (reconstructedOption === selectedEventOption) {
        clinicType = String(rowData[TYPE_COLUMN_INDEX - 1] || '').trim().toLowerCase();
        break;
      }
    }
    
    // Determine which response sheet to use
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentFormSheetName = e.range.getSheet().getName();
    let expectedSheetName;
    
    if (!clinicType) {
      // Event not found in Data Clinics - fall back to using the current sheet
      logMessage(`WARNING: Event "${selectedEventOption}" not found in Data Clinics sheet. Using current form sheet: "${currentFormSheetName}"`);
      expectedSheetName = currentFormSheetName;
      // Try to infer clinic type from sheet name
      if (currentFormSheetName === OPEN_FORM_RESPONSE_SHEET_NAME) {
        clinicType = 'open';
      } else if (currentFormSheetName === BESLOTEN_FORM_RESPONSE_SHEET_NAME) {
        clinicType = 'besloten';
      }
    } else {
      expectedSheetName = clinicType === 'open' ? OPEN_FORM_RESPONSE_SHEET_NAME : BESLOTEN_FORM_RESPONSE_SHEET_NAME;
      logMessage(`Form submission for ${clinicType} event "${selectedEventOption}" - current sheet: "${currentFormSheetName}", expected: "${expectedSheetName}"`);
    }
    
    // Now build duplicate lookup map ONLY for the relevant sheet (Open or Besloten)
    const existingParticipantsMap = {}; // Key: "email|eventname" -> { sheetName, rowIndex }
    const relevantSheet = ss.getSheetByName(expectedSheetName);
    
    // Get the current submission row number to exclude it from duplicate detection
    const currentSubmissionRow = e.range.getRow();
    const currentSubmissionSheetName = e.range.getSheet().getName();
    
    if (relevantSheet) {
      const data = relevantSheet.getDataRange().getValues();
      if (data.length >= 2) {
        const sheetHeaders = data.shift();
        const headerMap = {};
        sheetHeaders.forEach((h, i) => headerMap[h] = i);
        
        if (headerMap[FORM_EMAIL_QUESTION_TITLE] !== undefined && headerMap[FORM_EVENT_QUESTION_TITLE] !== undefined) {
          data.forEach((row, index) => {
            const rowNumber = index + 2; // +2 because: +1 for 1-based indexing, +1 for header row
            
            // Skip the current submission row to avoid detecting itself as a duplicate
            if (relevantSheet.getName() === currentSubmissionSheetName && rowNumber === currentSubmissionRow) {
              return; // Skip this row
            }
            
            const email = String(row[headerMap[FORM_EMAIL_QUESTION_TITLE]] || '').trim().toLowerCase();
            const eventName = String(row[headerMap[FORM_EVENT_QUESTION_TITLE]] || '').replace(/\s\(.*\)$/, '').trim();
            if (email && eventName) {
              const key = `${email}|${eventName}`;
              if (!existingParticipantsMap[key]) {
                existingParticipantsMap[key] = { sheetName: expectedSheetName, rowIndex: rowNumber };
              }
            }
          });
        }
      }
    }
    
    // Now process the event data (we already know clinicType from above)
    for (let i = 0; i < allData.length; i++) {
      const rowData = allData[i];
      const sheetDateValue = rowData[DATE_COLUMN_INDEX - 1];
      if (!sheetDateValue) continue;
      
      const reconstructedOption = `${getDutchDateString(new Date(sheetDateValue))} ${String(rowData[TIME_COLUMN_INDEX - 1]).trim()}, ${String(rowData[LOCATION_COLUMN_INDEX - 1]).trim()}`;
      
      if (reconstructedOption === selectedEventOption) {
        foundRowIndex = i + DATA_CLINICS_START_ROW;

        // Check if this participant (email + event) already exists in the CORRECT response sheet (Open or Besloten)
        const emailForLookup = String(placeholderMap['<Email>'] || '').trim().toLowerCase();
        const lookupKey = `${emailForLookup}|${selectedEventOption}`;
        const existingParticipantInfo = existingParticipantsMap[lookupKey];
        
        if (existingParticipantInfo) {
          logMessage(`DUPLICATE DETECTED: ${emailForLookup} already exists for event "${selectedEventOption}" in sheet "${existingParticipantInfo.sheetName}" at row ${existingParticipantInfo.rowIndex}. Updating existing participant.`);
          
          // Duplicate detected: update existing participant row with any new info and reuse their participant number/folder
          const updateSheet = ss.getSheetByName(existingParticipantInfo.sheetName);
          if (updateSheet) {
            const updateHeaders = updateSheet.getRange(1, 1, 1, updateSheet.getLastColumn()).getValues()[0];
            const updateHeaderMap = {};
            updateHeaders.forEach((h, idx) => updateHeaderMap[h] = idx);

            // Update fields if present in submission (update the existing row from Excel import)
            if (updateHeaderMap[FORM_FIRST_NAME_QUESTION_TITLE] !== undefined && placeholderMap['<Voornaam>']) {
              updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_FIRST_NAME_QUESTION_TITLE] + 1).setValue(placeholderMap['<Voornaam>']);
              logMessage(`  Updated Voornaam to: ${placeholderMap['<Voornaam>']}`);
            }
            if (updateHeaderMap[FORM_LAST_NAME_QUESTION_TITLE] !== undefined && placeholderMap['<Achternaam>']) {
              updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_LAST_NAME_QUESTION_TITLE] + 1).setValue(placeholderMap['<Achternaam>']);
              logMessage(`  Updated Achternaam to: ${placeholderMap['<Achternaam>']}`);
            }
            if (updateHeaderMap[FORM_PHONE_QUESTION_TITLE] !== undefined && placeholderMap['<Telefoonnummer>']) {
              updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_PHONE_QUESTION_TITLE] + 1).setValue(placeholderMap['<Telefoonnummer>']);
              logMessage(`  Updated Telefoonnummer to: ${placeholderMap['<Telefoonnummer>']}`);
            }
            if (updateHeaderMap[FORM_DOB_QUESTION_TITLE] !== undefined && placeholderMap['<Geboortedatum>']) {
              updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_DOB_QUESTION_TITLE] + 1).setValue(placeholderMap['<Geboortedatum>']);
              logMessage(`  Updated Geboortedatum to: ${placeholderMap['<Geboortedatum>']}`);
            }
            if (updateHeaderMap[FORM_CITY_QUESTION_TITLE] !== undefined && placeholderMap['<Woonplaats>']) {
              updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_CITY_QUESTION_TITLE] + 1).setValue(placeholderMap['<Woonplaats>']);
              logMessage(`  Updated Woonplaats to: ${placeholderMap['<Woonplaats>']}`);
            }
            if (updateHeaderMap[FORM_OPMERKING_QUESTION_TITLE] !== undefined && placeholderMap['<Opmerking>']) {
              updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_OPMERKING_QUESTION_TITLE] + 1).setValue(placeholderMap['<Opmerking>']);
              logMessage(`  Updated Opmerking to: ${placeholderMap['<Opmerking>']}`);
            }
            if (updateHeaderMap[FORM_MOTIVATIE_QUESTION_TITLE] !== undefined && placeholderMap['<Motivatie>']) {
              updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_MOTIVATIE_QUESTION_TITLE] + 1).setValue(placeholderMap['<Motivatie>']);
              logMessage(`  Updated Motivatie to: ${placeholderMap['<Motivatie>']}`);
            }
            
            // Update timestamp to reflect the latest submission
            if (updateHeaderMap['Timestamp'] !== undefined) {
              updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap['Timestamp'] + 1).setValue(new Date());
              logMessage(`  Updated Timestamp to current time`);
            }
            
            // Update registration method to show it came from both sources
            if (updateHeaderMap[FORM_REG_METHOD_QUESTION_TITLE] !== undefined) {
              const currentMethod = String(updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_REG_METHOD_QUESTION_TITLE] + 1).getValue() || '').trim();
              if (currentMethod === 'Excel Import') {
                updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[FORM_REG_METHOD_QUESTION_TITLE] + 1).setValue('Excel Import + Form');
                logMessage(`  Updated registration method to: Excel Import + Form`);
              }
            }

            // Retrieve existing participant number and folder ID
            if (updateHeaderMap[DEELNEMERNUMMER_HEADER] !== undefined) participantSequenceNumber = String(updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[DEELNEMERNUMMER_HEADER] + 1).getValue() || '').trim();
            if (updateHeaderMap[DRIVE_FOLDER_ID_HEADER] !== undefined) participantSubfolderId = String(updateSheet.getRange(existingParticipantInfo.rowIndex, updateHeaderMap[DRIVE_FOLDER_ID_HEADER] + 1).getValue() || '').trim();
            
            // Rename participant folder if first name or last name was updated
            if (participantSubfolderId && (placeholderMap['<Voornaam>'] || placeholderMap['<Achternaam>'])) {
              try {
                const folder = DriveApp.getFolderById(participantSubfolderId);
                const formattedParticipantNumber = Utilities.formatString('%02d', parseInt(participantSequenceNumber) || 0);
                const newFolderName = `${formattedParticipantNumber} ${placeholderMap['<Voornaam>']} ${placeholderMap['<Achternaam>']}`.replace(/\s+/g, ' ').trim();
                const currentFolderName = folder.getName();
                
                if (currentFolderName !== newFolderName) {
                  folder.setName(newFolderName);
                  logMessage(`  Participant folder renamed from "${currentFolderName}" to "${newFolderName}"`);
                }
              } catch (folderError) {
                logMessage(`  WAARSCHUWING: Kon participant folder niet hernoemen. Fout: ${folderError.message}`);
              }
            }
            
            SpreadsheetApp.flush(); // Ensure updates are committed
          }

          // We do NOT increment booked seats or create new folders for duplicates
          placeholderMap['<Deelnemernummer>'] = participantSequenceNumber;
          isDuplicate = true; // Mark as duplicate
          break;
        }

        // No duplicate found: proceed to add as new participant (increase booked seats, create folder)
        const currentBookedSeats = (rowData[BOOKED_SEATS_COLUMN_INDEX - 1] || 0);
        const participantEmail = placeholderMap['<Email>'];
        const isNonParticipant = isNonParticipantEmail(participantEmail);
        
        if (isNonParticipant) {
          // Non-participant (test/host account): do NOT increment booked seats count
          logMessage(`Non-participant account gedetecteerd (${participantEmail}). Deelnemertelling NIET verhoogd.`);
          participantSequenceNumber = 'xx';
        } else {
          // Regular participant: increment the count
          const newBookedSeats = currentBookedSeats + 1;
          dataClinicsSheet.getRange(foundRowIndex, BOOKED_SEATS_COLUMN_INDEX).setValue(newBookedSeats);
          participantSequenceNumber = Utilities.formatString('%02d', newBookedSeats);
        }
        placeholderMap['<Deelnemernummer>'] = participantSequenceNumber;

        const eventDateFormatted = Utilities.formatDate(new Date(sheetDateValue), FORMATTING_TIME_ZONE, DATE_FORMAT_YYYYMMDD);
        const timeFormatted = String(rowData[TIME_COLUMN_INDEX - 1]).trim().replace(/:|\.\|/g, '');
        const eventFolderName = `${eventDateFormatted} ${timeFormatted} ${String(rowData[LOCATION_COLUMN_INDEX-1]).trim()}`;
        
        const parentFolder = DriveApp.getFolderById(PARENT_EVENT_FOLDER_ID);
        const folders = parentFolder.getFoldersByName(eventFolderName);
        
        // Check for duplicate folders
        let eventFolder;
        if (folders.hasNext()) {
          eventFolder = folders.next();
          if (folders.hasNext()) {
            logMessage(`WAARSCHUWING: Meerdere mappen gevonden met naam "${eventFolderName}". Eerste map wordt gebruikt (ID: ${eventFolder.getId()}).`);
          }
        } else {
          eventFolder = parentFolder.createFolder(eventFolderName);
          logMessage(`Nieuwe Event Folder aangemaakt: "${eventFolderName}" (ID: ${eventFolder.getId()}).`);
        }

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
    
    // If this is a duplicate, delete the new form submission row since we updated the existing Excel import row
    if (isDuplicate) {
      logMessage(`Deleting duplicate form submission row ${targetRow} in sheet "${targetSheet.getName()}".`);
      targetSheet.deleteRow(targetRow);
      logMessage(`Duplicate row deleted. Existing participant row was updated instead.`);
    } else {
      // Normal flow: populate the new form submission row with participant number and folder
      if (pNumCol !== -1) targetSheet.getRange(targetRow, pNumCol + 1).setValue(participantSequenceNumber);
      if (fIdCol !== -1) targetSheet.getRange(targetRow, fIdCol + 1).setValue(participantSubfolderId);
    }

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
      
      const mailOptions = { name: mergedMail.senderName, htmlBody: mergedMail.htmlBody };
      if (fromAlias) {
        mailOptions.from = fromAlias;
      }
      GmailApp.sendEmail(placeholderMap['<Email>'], mergedMail.subject, '', mailOptions);
      
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