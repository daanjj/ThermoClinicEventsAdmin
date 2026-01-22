// MailMerge v2.4 - Optimized for Performance and Robustness

// This file contains all functions related to the mail merge feature,
// including showing the dialog, retrieving data, merging templates, and sending emails.

function showMailMergeDialog() {
  // Check Gmail alias before showing dialog - we can show UI alerts here (menu context)
  const desiredAlias = EMAIL_SENDER_ALIAS;
  const availableAliases = GmailApp.getAliases();
  
  if (!availableAliases.includes(desiredAlias)) {
    const currentUser = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'onbekend';
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Verkeerd account',
      `Let op: Email alias "${desiredAlias}" niet gevonden.\n\n` +
      `Je bent ingelogd als: ${currentUser}\n` +
      `Voor correcte afzender moet je inloggen als: joost@thermoclinics.nl\n\n` +
      `Als je doorgaat worden de mails verstuurd vanuit ${currentUser}.\n\n` +
      `Wil je doorgaan?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      logMessage(`Mail merge geannuleerd door gebruiker. Ingelogd als ${currentUser}, alias ${desiredAlias} niet beschikbaar.`);
      return;
    }
    logMessage(`WAARSCHUWING: Gebruiker ${currentUser} gaat door met mail merge zonder alias ${desiredAlias}.`);
  }
  
  const htmlTemplate = HtmlService.createTemplateFromFile('MailMergeDialog');
  htmlTemplate.clinics = getAvailableClinicsList();
  htmlTemplate.templates = getMailTemplates();
  htmlTemplate.genericAttachments = getGenericAttachments();
  SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setWidth(400).setHeight(550), 'Mail Merge voor Clinic');
}

/**
 * [OPTIMIZED]
 * Prepares a mail template by performing the slow operations (copying, fetching HTML) once.
 * @param {string} templateId The ID of the Google Doc template.
 * @returns {object|null} An object with {rawHtmlBody, subjectTemplate, senderName} or null on failure.
 */
function prepareMailTemplate(templateId) {
  let tempDocFile = null;
  try {
    tempDocFile = DriveApp.getFileById(templateId).makeCopy(`[TEMP] Prep for Merge - ${new Date().getTime()}`);
    const tempDoc = DocumentApp.openById(tempDocFile.getId());
    const fullText = tempDoc.getBody().getText();
    const lines = fullText.split('\n');
    
    let senderName = FALLBACK_EMAIL_SENDER_NAME;
    if (lines[0] && lines[0].trim().toLowerCase().startsWith('van:')) {
      senderName = lines[0].substring(lines[0].indexOf(':') + 1).trim();
    }
    let subjectTemplate = `Informatie over: <Eventnaam>`;
    if (lines.length > 1 && lines[1].trim().toLowerCase().startsWith('onderwerp:')) {
      subjectTemplate = lines[1].substring(lines[1].indexOf(':') + 1).trim();
    }

    const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${tempDocFile.getId()}&exportFormat=html`;
    const params = { method: "get", headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
    const exportedHtml = UrlFetchApp.fetch(url, params).getContentText();

    let bodyContent = '';
    const bodyMatch = exportedHtml.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    bodyContent = (bodyMatch && bodyMatch[1]) ? bodyMatch[1] : exportedHtml;

    const allParagraphs = bodyContent.match(/<(p|h[1-6])[^>]*>[\s\S]*?<\/(p|h[1-6])>/gi) || [];
    const mainContentParagraphs = allParagraphs.slice(3);
    let rawHtmlBody = mainContentParagraphs.join('\n');
    
    const lt_from_code = String.fromCharCode(38, 108, 116, 59); // &lt;
    const gt_from_code = String.fromCharCode(38, 103, 116, 59); // &gt;
    rawHtmlBody = rawHtmlBody.split(lt_from_code).join('<').split(gt_from_code).join('>');
    
    return { rawHtmlBody, subjectTemplate, senderName };
  } catch (e) {
    Logger.log(`CRITICAL ERROR in prepareMailTemplate (template ID: ${templateId}): ${e.toString()}\n${e.stack}`);
    return null;
  } finally {
    if (tempDocFile) {
      tempDocFile.setTrashed(true);
    }
  }
}

// ADD THIS NEW FUNCTION TO MailMerge v2.4.js

/**
 * Merges a single template for a specific recipient.
 * This is a replacement for the old mergeTemplateInDoc for non-batch operations.
 * It uses the robust prepareMailTemplate function internally.
 * @param {string} templateId The ID of the Google Doc template.
 * @param {object} placeholderMap A map of placeholders to their values.
 * @returns {object} An object with { htmlBody, subject, senderName }.
 */
function mergeSingleTemplate(templateId, placeholderMap) {
  // 1. Prepare the template (this does the slow work)
  const preparedTemplate = prepareMailTemplate(templateId);
  if (!preparedTemplate) {
    throw new Error(`Failed to prepare template with ID ${templateId}.`);
  }

  // 2. Perform a fast merge using the prepared content
  let participantBodyHtml = resolveTimeArithmeticPlaceholders(preparedTemplate.rawHtmlBody, placeholderMap);
  let participantSubject = preparedTemplate.subjectTemplate;

  for (const placeholder in placeholderMap) {
    const value = String(placeholderMap[placeholder] || '');
    participantBodyHtml = participantBodyHtml.replaceAll(placeholder, value);
    participantSubject = participantSubject.replaceAll(placeholder, value);
  }
  
  // Clean up any double spaces or space before punctuation in subject
  participantSubject = participantSubject.replace(/\s+([!?.,;:])/g, '$1').replace(/\s+/g, ' ').trim();

  // 3. Perform final cleaning and formatting
  const finalParagraphs = participantBodyHtml.match(/<(p|h[1-6])[^>]*>[\s\S]*?<\/(p|h[1-6])>/gi) || [];
  const finalCleanedElements = finalParagraphs.map(p => {
    const textContent = p.replace(/<[^>]+>/g, ' ').trim();
    if (textContent === '') return null; // Mark empty paragraphs for removal
    if (/^\d+:/.test(textContent)) {
      return p.replace(/<p/i, '<p style="padding-left: 25px;"');
    }
    return p;
  }).filter(Boolean); // Remove nulls
  const finalBodyContent = finalCleanedElements.join('\n');

  // 4. Construct the final HTML email structure
  const finalHtml = `
    <!DOCTYPE html><html lang="nl"><head><meta charset="UTF-8"><title>${participantSubject}</title><style>/* ... your full CSS styles ... */</style></head><body style="background-color: #f4f4f4; margin: 0 !important; padding: 0 !important;"><table border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 800px;"><tr><td align="center" style="padding: 20px 0;"><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td align="left" bgcolor="#ffffff" style="padding: 25px 40px; font-family: sans-serif; font-size: 16px; color: #333;">${finalBodyContent}</td></tr></table></td></tr></table></body></html>`;

  return {
    htmlBody: finalHtml,
    subject: participantSubject,
    senderName: preparedTemplate.senderName
  };
}


/**
 * [REWRITTEN & BULLETPROOF]
 * Performs the mail merge with robust error handling for individual rows.
 * This will not fail on bad data and will report which row caused an issue.
 */
function performMailMerge(selectedClinic, selectedTemplateId, selectedTemplateName, selectedAttachmentIds) {
  const logHeader = `----- START mailmerge voor ${selectedClinic} met sjabloon ${selectedTemplateName} -----`;
  logMessage(logHeader);
  
  // Verify Gmail alias is available - script should run under joost@thermoclinics.nl
  // Note: Cannot show UI dialogs when called from HTML dialog, so we log warning and proceed
  const desiredAlias = EMAIL_SENDER_ALIAS;
  const availableAliases = GmailApp.getAliases();
  let fromAlias = availableAliases.includes(desiredAlias) ? desiredAlias : null;
  
  if (!fromAlias) {
    const currentUser = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'onbekend';
    logMessage(`WAARSCHUWING: Email alias "${desiredAlias}" niet gevonden. Ingelogd als ${currentUser}. Voor correcte afzender moet je inloggen als joost@thermoclinics.nl. Emails worden verstuurd vanuit ${currentUser}.`);
  }

  try {
    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const dataClinicsSheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
    const allDataClinicsData = dataClinicsSheet.getDataRange().getValues();
    
    let clinicType = '', clinicFullTime = '';
    for (const row of allDataClinicsData) {
      const dateValue = row[DATE_COLUMN_INDEX - 1];
      if (!dateValue) continue;
      const timeText = String(row[TIME_COLUMN_INDEX - 1] || '').trim();
      const locationText = String(row[LOCATION_COLUMN_INDEX - 1] || '').trim();
      const typeValue = String(row[TYPE_COLUMN_INDEX - 1] || '').trim();
      const reconstructedOption = `${getDutchDateString(new Date(dateValue))} ${timeText}, ${locationText}`;
      if (reconstructedOption === selectedClinic) {
        clinicType = typeValue; clinicFullTime = timeText; break;
      }
    }
    if (!clinicType) throw new Error(`Could not find data for clinic: "${selectedClinic}".`);

    const responseSheetName = clinicType.toLowerCase() === 'open' ? OPEN_FORM_RESPONSE_SHEET_NAME : BESLOTEN_FORM_RESPONSE_SHEET_NAME;
    const responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(responseSheetName);
    if (!responseSheet) throw new Error(`Response sheet "${responseSheetName}" not found.`);
    
    const responseDataWithHeaders = responseSheet.getDataRange().getValues();
    const headers = responseDataWithHeaders.shift();
    const headerMap = headers.reduce((acc, header, index) => { acc[String(header).trim()] = index; return acc; }, {});
    
    const participantsWithIndex = responseDataWithHeaders
      .map((row, index) => ({ row, originalIndex: index + 2 })) // +2 for 1-based index and header row
      .filter(item => (item.row[headerMap[FORM_EVENT_QUESTION_TITLE]] || '').replace(/\s\(.*\)$/, '').trim() === selectedClinic);

    if (participantsWithIndex.length === 0) {
      logMessage(`Geen deelnemers gevonden voor de clinic: "${selectedClinic}".`);
      throw new Error(`Geen deelnemers gevonden voor de clinic: "${selectedClinic}".`);
    }

    const preparedTemplate = prepareMailTemplate(selectedTemplateId);
    if (!preparedTemplate) throw new Error(`Failed to prepare mail template. Check logs.`);

    const genericAttachments = (selectedAttachmentIds || []).map(id => {
      try { return DriveApp.getFileById(id).getBlob(); } catch (e) { logMessage(`WAARSCHUWING: Kon generieke bijlage ID ${id} niet ophalen.`); return null; }
    }).filter(Boolean);
    const genericAttachmentNames = genericAttachments.map(blob => blob.getName());

    let sentCount = 0;
    const isParticipantAttachmentTemplate = selectedTemplateName.toLowerCase().includes('bijlage');
    const driveFolderIdColumnIndex = headerMap[DRIVE_FOLDER_ID_HEADER];

    for (const item of participantsWithIndex) {
      const { row: participantRow, originalIndex } = item;
      let participantIdentifier = `rij ${originalIndex} in sheet '${responseSheetName}'`;

      try {
        const email = String(participantRow[headerMap[FORM_EMAIL_QUESTION_TITLE]] || '').trim();
        if (email) participantIdentifier = email;

        if (!email) {
          logMessage(`Overgeslagen: Deelnemer op ${participantIdentifier} heeft geen e-mailadres.`);
          continue;
        }

        const placeholderMap = {
          '<Voornaam>': String(participantRow[headerMap[FORM_FIRST_NAME_QUESTION_TITLE]] || ''),
          '<Achternaam>': String(participantRow[headerMap[FORM_LAST_NAME_QUESTION_TITLE]] || ''),
          '<Email>': email,
          '<Telefoonnummer>': String(participantRow[headerMap[FORM_PHONE_QUESTION_TITLE]] || ''),
          '<Geboortedatum>': String(participantRow[headerMap[FORM_DOB_QUESTION_TITLE]] || ''),
          '<Woonplaats>': String(participantRow[headerMap[FORM_CITY_QUESTION_TITLE]] || ''),
          '<Deelnemernummer>': String(participantRow[headerMap[DEELNEMERNUMMER_HEADER]] || ''),
          '<CORE-mailadres>': String(participantRow[headerMap[FORM_CORE_MAIL_HEADER]] || ''),
          '<Eventnaam>': selectedClinic
        };
        const eventParts = selectedClinic.split(',');
        placeholderMap['<Locatie>'] = eventParts.length > 1 ? eventParts[1].trim() : '';
        const dateTimePart = eventParts[0].trim();
        const timeStartIndex = dateTimePart.indexOf(clinicFullTime);
        placeholderMap['<Datum>'] = (timeStartIndex > -1) ? dateTimePart.substring(0, timeStartIndex).trim() : dateTimePart.replace(clinicFullTime, '').trim();
        placeholderMap['<Tijd>'] = clinicFullTime;
        placeholderMap['<Starttijd>'] = clinicFullTime.split('-')[0].trim();
        if (placeholderMap['<Geboortedatum>']) { try { placeholderMap['<Geboortedatum>'] = getDutchDateString(new Date(placeholderMap['<Geboortedatum>'])); } catch (e) { /* ignore formatting error */ } }

        let participantBodyHtml = resolveTimeArithmeticPlaceholders(preparedTemplate.rawHtmlBody, placeholderMap);
        let participantSubject = preparedTemplate.subjectTemplate;

        for (const placeholder in placeholderMap) {
          const value = String(placeholderMap[placeholder] || '');
          participantBodyHtml = participantBodyHtml.replaceAll(placeholder, value);
          participantSubject = participantSubject.replaceAll(placeholder, value);
        }
        
        // Clean up any double spaces or space before punctuation in subject
        participantSubject = participantSubject.replace(/\s+([!?.,;:])/g, '$1').replace(/\s+/g, ' ').trim();

        const finalParagraphs = participantBodyHtml.match(/<(p|h[1-6])[^>]*>[\s\S]*?<\/(p|h[1-6])>/gi) || [];
        const finalCleanedElements = finalParagraphs.map(p => {
            const textContent = p.replace(/<[^>]+>/g, ' ').trim();
            if (textContent === '') return null;
            if (/^\d+:/.test(textContent)) return p.replace(/<p/i, '<p style="padding-left: 25px;"');
            return p;
        }).filter(Boolean);
        const finalBodyContent = finalCleanedElements.join('\n');
        const finalHtml = `<!DOCTYPE html>...${finalBodyContent}...</html>`; // (Your full HTML boilerplate here)
        
        // START INSERTED CODE

        // Get the names of generic attachments once per participant for logging
        const genericAttachmentNames = genericAttachments.map(blob => blob.getName());
        const hasGenericAttachments = genericAttachments.length > 0;

        if (isParticipantAttachmentTemplate) {
          // --- LOGIC FOR ATTACHMENT TEMPLATES ---
          const participantFolderId = String(participantRow[driveFolderIdColumnIndex] || '').trim();
          const participantFilesToSend = []; // Holds File objects for moving
          const participantBlobs = [];       // Holds Blobs for attaching

          if (participantFolderId) {
            try {
              const folder = DriveApp.getFolderById(participantFolderId);
              const filesIterator = folder.getFiles();
              while (filesIterator.hasNext()) {
                const file = filesIterator.next();
                participantFilesToSend.push(file); // Store the File object
                participantBlobs.push(file.getBlob()); // Store the Blob for the email
              }
            } catch (attachError) {
              logMessage(`WAARSCHUWING: Kon specifieke bijlagen niet ophalen voor ${participantIdentifier} (map ID: ${participantFolderId}). Fout: ${attachError.message}`);
              Logger.log(`Participant Attachment Error for ${participantIdentifier} [FolderID: ${participantFolderId}]: ${attachError.toString()}`);
            }
          } else {
              logMessage(`Info: Geen map-ID gevonden voor ${participantIdentifier}. Alleen generieke bijlagen worden verstuurd (indien geselecteerd).`);
          }
          
          if (participantBlobs.length > 0 || hasGenericAttachments) {
            const mailOptions = { 
              name: preparedTemplate.senderName, 
              htmlBody: finalHtml,
              attachments: [...genericAttachments, ...participantBlobs]
            };
            if (fromAlias) {
              mailOptions.from = fromAlias;
            }
            
            GmailApp.sendEmail(email, participantSubject, '', mailOptions);
            sentCount++;
            
            // --- START: Enhanced Logging with Filenames ---
            const specificAttachmentNames = participantFilesToSend.map(file => file.getName());
            const allAttachmentNames = [...genericAttachmentNames, ...specificAttachmentNames];
            
            let attachmentsLogString = '';
            if (allAttachmentNames.length > 0) {
              attachmentsLogString = ` [Bijlagen: ${allAttachmentNames.join(', ')}]`;
            }

            logMessage(`Mail verstuurd aan: ${participantIdentifier}, Onderwerp: "${participantSubject}"${attachmentsLogString}`);
            // --- END: Enhanced Logging with Filenames ---

            if (participantFilesToSend.length > 0) {
              try {
                const folder = DriveApp.getFolderById(participantFolderId); 
                const subfolderName = 'Reeds verstuurde bijlagen';
                let destinationFolder;
                const existingFolders = folder.getFoldersByName(subfolderName);

                if (existingFolders.hasNext()) {
                  destinationFolder = existingFolders.next();
                } else {
                  destinationFolder = folder.createFolder(subfolderName);
                  // logMessage(`Info: Submap '${subfolderName}' aangemaakt in map ${participantFolderId}.`);
                }

                for (const fileToMove of participantFilesToSend) {
                  fileToMove.moveTo(destinationFolder);
                }
                // logMessage(`Info: ${participantFilesToSend.length} specifieke bijlage(n) succesvol verplaatst naar '${subfolderName}' voor ${participantIdentifier}.`);

              } catch (moveError) {
                logMessage(`FOUT: Kon bestanden niet verplaatsen voor ${participantIdentifier} na verzending. Fout: ${moveError.message}`);
                Logger.log(`File Move Error for ${participantIdentifier} [FolderID: ${participantFolderId}]: ${moveError.toString()}`);
              }
            }

          } else {
            logMessage(`Overgeslagen: ${participantIdentifier} - Geen generieke of specifieke bijlagen gevonden om te versturen.`);
            continue;
          }
        
        } else {
          // --- LOGIC FOR NON-ATTACHMENT TEMPLATES ---
          const mailOptions = { name: preparedTemplate.senderName, htmlBody: finalHtml, attachments: [...genericAttachments] };
          if (fromAlias) {
            mailOptions.from = fromAlias;
          }
          GmailApp.sendEmail(email, participantSubject, '', mailOptions);
          sentCount++;
          
          // --- START: Enhanced Logging with Filenames ---
          let attachmentsLogString = '';
          if (genericAttachmentNames.length > 0) {
            attachmentsLogString = ` [Bijlagen: ${genericAttachmentNames.join(', ')}]`;
          }
          logMessage(`Mail verstuurd aan: ${participantIdentifier}, Onderwerp: "${participantSubject}"${attachmentsLogString}`);
          // --- END: Enhanced Logging with Filenames ---
        }

        // END INSERTED CODE

      } catch (err) {
        logMessage(`FOUT bij verwerken van ${participantIdentifier}: ${err.message}`);
        Logger.log(`Mail Merge Participant Error for ${participantIdentifier}: ${err.toString()}\n${err.stack}`);
      }
    }
    
    const successMessage = `Mail merge voltooid! ${sentCount} mail(s) verstuurd.`;
    logMessage(`----- EINDE mailmerge: ${sentCount} mail(s) verstuurd -----`);
    return successMessage; // Return success message for HTML dialog to display

  } catch (err) {
    Logger.log(`performMailMerge FATAL ERROR: ${err.toString()}\n${err.stack}`);
    logMessage(`Mailmerge FATALE FOUT: ${err.message}`);
    throw err; // Re-throw for HTML dialog to handle
  } finally {
    flushLogs();
  }
}


function getAvailableClinicsList() {
  try {
    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const sheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_CLINICS_START_ROW) return [];
    
    // Read only relevant columns for performance
    const maxColumnIndexToRead = Math.max(DATE_COLUMN_INDEX, TIME_COLUMN_INDEX, LOCATION_COLUMN_INDEX, BOOKED_SEATS_COLUMN_INDEX);
    const allData = sheet.getRange(DATA_CLINICS_START_ROW, 1, lastRow - DATA_CLINICS_START_ROW + 1, maxColumnIndexToRead).getValues();
    
    const clinicOptions = [];
    for (let i = 0; i < allData.length; i++) {
      const rowData = allData[i];
      const dateValue = rowData[DATE_COLUMN_INDEX - 1];
      const timeText = rowData[TIME_COLUMN_INDEX - 1];
      const locationText = rowData[LOCATION_COLUMN_INDEX - 1];
      const bookedSeatsRaw = rowData[BOOKED_SEATS_COLUMN_INDEX - 1];
      
      const numBookedSeats = (bookedSeatsRaw === '' || bookedSeatsRaw === null || bookedSeatsRaw === undefined) ? 0 : parseInt(bookedSeatsRaw, 10);
      
      // Only include clinics with at least one participant
      if (isNaN(numBookedSeats) || numBookedSeats === 0) {
        continue;
      }
      
      let actualDateObject = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
      if (isNaN(actualDateObject.getTime())) {
        continue; // Skip invalid dates
      }
      
      const formattedDatePart = getDutchDateString(actualDateObject);
      const combinedOption = `${formattedDatePart} ${String(timeText).trim()}, ${String(locationText).trim()}`;
      clinicOptions.push(combinedOption);
    }
    return clinicOptions;
  } catch (err) {
    Logger.log(`getAvailableClinicsList ERROR: ${err.toString()}`);
    logMessage(`FOUT bij ophalen beschikbare clinics voor mailmerge: ${err.message}`);
    return [];
  }
}

function getMailTemplates() {
  try {
    const folder = DriveApp.getFolderById(MAIL_TEMPLATE_FOLDER_ID);
    // Search for Google Docs files with "mailsjabloon" in their title
    const files = folder.searchFiles('title contains "mailsjabloon" and mimeType = "application/vnd.google-apps.document"');
    const templates = [];
    while (files.hasNext()) {
      const file = files.next();
      templates.push({
        name: file.getName(),
        id: file.getId()
      });
    }
    return templates;
  } catch (err) {
    Logger.log(`getMailTemplates ERROR: ${err.toString()}`);
    logMessage(`FOUT bij ophalen mailsjablonen: ${err.message}`);
    return [];
  }
}

function getGenericAttachments() {
  const attachments = [];
  try {
    const genericFolder = DriveApp.getFolderById(GENERIC_ATTACHMENTS_FOLDER_ID);
    const files = genericFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      attachments.push({
        name: file.getName(),
        id: file.getId()
      });
    }
  } catch (e) {
    Logger.log(`Could not get generic attachments from folder ID: ${GENERIC_ATTACHMENTS_FOLDER_ID}. Error: ${e.toString()}`);
    logMessage(`WAARSCHUWING: Kon generieke bijlagen niet ophalen uit map ID ${GENERIC_ATTACHMENTS_FOLDER_ID}. Fout: ${e.message}`);
  }
  return attachments;
}