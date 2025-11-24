// This file manages functions specifically related to the CORE app integration,
// including reminders and updating CORE app email addresses in the response sheets.

function showCoreReminderDialog() {
  const clinics = getClinicsForReminder();
  const htmlTemplate = HtmlService.createTemplateFromFile('COREReminderDialog');
  htmlTemplate.clinics = clinics;
  const htmlOutput = htmlTemplate.evaluate().setWidth(400).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Stuur reminder om CORE-app te installeren');
}

function getClinicsForReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const openResponsesSheet = ss.getSheetByName(OPEN_FORM_RESPONSE_SHEET_NAME);
  const beslotenResponsesSheet = ss.getSheetByName(BESLOTEN_FORM_RESPONSE_SHEET_NAME);
  const clinicsWithMissingMail = new Set();

  const findMissingEmails = (sheet) => {
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const eventColIdx = headers.indexOf(FORM_EVENT_QUESTION_TITLE);
    const coreMailColIdx = headers.indexOf(FORM_CORE_MAIL_HEADER);
    if (eventColIdx === -1 || coreMailColIdx === -1) {
      logMessage(`WAARSCHUWING: Kolom '${FORM_EVENT_QUESTION_TITLE}' of '${FORM_CORE_MAIL_HEADER}' niet gevonden in tabblad '${sheet.getName()}'.`);
      return;
    }
    data.forEach(row => {
      const coreMail = row[coreMailColIdx] || '';
      const eventName = (row[eventColIdx] || '').replace(/\s\(.*\)$/, '').trim();
      if (eventName && coreMail.trim() === '') {
        clinicsWithMissingMail.add(eventName);
      }
    });
  };

  findMissingEmails(openResponsesSheet);
  findMissingEmails(beslotenResponsesSheet);

  if (clinicsWithMissingMail.size === 0) return [];

  const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
  const sheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
  if (!sheet) {
    logMessage(`FOUT: Sheet '${DATA_CLINICS_SHEET_NAME}' niet gevonden voor CORE-app reminder.`);
    return [];
  }

  const allData = sheet.getRange(DATA_CLINICS_START_ROW, 1, sheet.getLastRow() - DATA_CLINICS_START_ROW + 1, LOCATION_COLUMN_INDEX).getValues();
  const finalClinicOptions = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  allData.forEach(rowData => {
    const dateValue = rowData[DATE_COLUMN_INDEX - 1];
    if (!dateValue) return;
    let eventDate = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
    if (isNaN(eventDate.getTime()) || eventDate < today) return; // Only future clinics

    const timeText = rowData[TIME_COLUMN_INDEX - 1];
    const locationText = rowData[LOCATION_COLUMN_INDEX - 1];
    const combinedOption = `${getDutchDateString(eventDate)} ${String(timeText).trim()}, ${String(locationText).trim()}`;

    if (clinicsWithMissingMail.has(combinedOption)) {
      finalClinicOptions.push(combinedOption);
    }
  });
  return finalClinicOptions;
}

function sendCoreAppReminder(selectedClinic) {
  const logHeader = `----- START CORE-app reminder voor ${selectedClinic} -----`;
  logMessage(logHeader);
  
  // Verify Gmail alias is available
  const desiredAlias = "info@thermoclinics.nl";
  const availableAliases = GmailApp.getAliases();
  const fromAlias = availableAliases.includes(desiredAlias) ? desiredAlias : null;
  
  if (!fromAlias) {
    logMessage(`WAARSCHUWING: Alias "${desiredAlias}" niet gevonden voor CORE reminder. Beschikbare aliassen: ${availableAliases.join(', ')}. Emails worden verstuurd zonder 'from' alias.`);
  }

  try {
    // ... (Logic to find clinic type and get participants remains the same)
    const dataClinicsSpreadsheet = SpreadsheetApp.openById(DATA_CLINICS_SPREADSHEET_ID);
    const dataClinicsSheet = dataClinicsSpreadsheet.getSheetByName(DATA_CLINICS_SHEET_NAME);
    const allDataClinicsData = dataClinicsSheet.getDataRange().getValues();
    let clinicType = '';
    for (const row of allDataClinicsData) {
      const dateValue = row[DATE_COLUMN_INDEX - 1];
      if (!dateValue) continue;
      const reconstructedOption = `${getDutchDateString(new Date(dateValue))} ${String(row[TIME_COLUMN_INDEX - 1]).trim()}, ${String(row[LOCATION_COLUMN_INDEX - 1]).trim()}`;
      if (reconstructedOption === selectedClinic) {
        clinicType = String(row[TYPE_COLUMN_INDEX - 1]).trim();
        break;
      }
    }
    if (!clinicType) throw new Error(`Kon type voor clinic "${selectedClinic}" niet vinden.`);

    const responseSheetName = clinicType.toLowerCase() === 'open' ? OPEN_FORM_RESPONSE_SHEET_NAME : BESLOTEN_FORM_RESPONSE_SHEET_NAME;
    const responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(responseSheetName);
    if (!responseSheet) throw new Error(`Response sheet "${responseSheetName}" niet gevonden.`);

    const responseData = responseSheet.getDataRange().getValues();
    const headers = responseData.shift();
    const headerMap = headers.reduce((acc, h, i) => { acc[String(h).trim()] = i; return acc; }, {});
    
    const participantsToRemind = responseData.filter(row => {
      const eventName = (row[headerMap[FORM_EVENT_QUESTION_TITLE]] || '').replace(/\s\(.*\)$/, '').trim();
      const coreMail = row[headerMap[FORM_CORE_MAIL_HEADER]] || '';
      return eventName === selectedClinic && coreMail.trim() === '';
    });

    if (participantsToRemind.length === 0) {
      logMessage(`Geen deelnemers gevonden voor clinic "${selectedClinic}" die een reminder nodig hebben.`);
      SpreadsheetApp.getUi().alert(`Alle deelnemers voor deze clinic hebben al een CORE-mailadres ingevuld.`);
      return;
    }

    // ===== OPTIMIZATION CHANGE IS HERE =====
    // 1. Prepare the template ONCE, before the loop starts.
    const preparedTemplate = prepareMailTemplate(CORE_APP_REMINDER_TEMPLATE_ID);
    if (!preparedTemplate) {
      throw new Error(`Kon de CORE-app reminder template niet voorbereiden.`);
    }
    // =======================================

    const templateName = DriveApp.getFileById(CORE_APP_REMINDER_TEMPLATE_ID).getName();
    let sentCount = 0;

    for (const participantRow of participantsToRemind) {
      try {
        const placeholderMap = {
          '<Voornaam>': String(participantRow[headerMap[FORM_FIRST_NAME_QUESTION_TITLE]] || '').trim(),
          '<Achternaam>': String(participantRow[headerMap[FORM_LAST_NAME_QUESTION_TITLE]] || '').trim(),
          '<Email>': String(participantRow[headerMap[FORM_EMAIL_QUESTION_TITLE]] || '').trim(),
          '<Eventnaam>': selectedClinic
        };
        const eventParts = selectedClinic.split(',');
        const dateTimePart = eventParts[0];
        const lastSpaceIndex = dateTimePart.lastIndexOf(' ');
        placeholderMap['<Tijd>'] = lastSpaceIndex > -1 ? dateTimePart.substring(lastSpaceIndex + 1).trim() : '';

        const recipientEmail = placeholderMap['<Email>'];
        if (!recipientEmail) {
          logMessage(`Overgeslagen: Deelnemer ${placeholderMap['<Voornaam>']} heeft geen e-mailadres.`);
          continue;
        }

        // ===== OPTIMIZATION CHANGE IS HERE =====
        // 2. Perform a fast merge using the prepared template content.
        // NO MORE CALL to mergeTemplateInDoc() here.
        let finalSubject = preparedTemplate.subjectTemplate;
        let finalHtmlBody = preparedTemplate.rawHtmlBody;

        for (const placeholder in placeholderMap) {
          const value = String(placeholderMap[placeholder] || '');
          finalSubject = finalSubject.replaceAll(placeholder, value);
          finalHtmlBody = finalHtmlBody.replaceAll(placeholder, value);
        }
        
        // Construct the final HTML (can be simplified if no complex cleaning is needed)
        const finalHtml = `<!DOCTYPE html>...${finalHtmlBody}...</html>`; // (Your full HTML boilerplate here)
        // =======================================

        const mailOptions = {
          name: preparedTemplate.senderName,
          htmlBody: finalHtml
        };
        if (fromAlias) {
          mailOptions.from = fromAlias;
        }
        GmailApp.sendEmail(recipientEmail, finalSubject, '', mailOptions);
        sentCount++;
        
        logMessage(`CORE-app Reminder verstuurd aan: ${recipientEmail}, Onderwerp: "${finalSubject}"`);

      } catch (mailError) {
        logMessage(`FOUT bij sturen reminder aan ${String(participantRow[headerMap[FORM_EMAIL_QUESTION_TITLE]] || 'onbekend')}: ${mailError.message}`);
      }
    }

    const mailText = (sentCount === 1) ? 'reminder' : 'reminders';
    const successMessage = `Versturen voltooid! ${sentCount} ${mailText} verstuurd.`;
    logMessage(`----- EINDE CORE-app reminder: ${sentCount} ${mailText} verstuurd -----`);
    SpreadsheetApp.getUi().alert(successMessage);

  } catch (err) {
    Logger.log(`sendCoreAppReminder CRITICAL ERROR: ${err.toString()}\n${err.stack}`);
    logMessage(`CORE-app reminder FOUT: ${err.message}`);
    SpreadsheetApp.getUi().alert(`Er is een fout opgetreden: ${err.message}`);
  } finally {
    flushLogs();
  }
}

function handleCoreAppFormSubmit(e) {
  if (!e.namedValues || !e.namedValues[CORE_APP_CLINIC_HEADER] || !e.namedValues[CORE_APP_EMAIL_CLINIC_HEADER]) {
    Logger.log(`FOUT: De data van het CORE-app formulier kon niet worden gelezen, ontbrekende velden.`);
    logMessage(`handleCoreAppFormSubmit ERROR: Ontbrekende velden in formulierinzending.`);
    return;
  }
  const clinicName = e.namedValues[CORE_APP_CLINIC_HEADER][0];
  const clinicEmail = e.namedValues[CORE_APP_EMAIL_CLINIC_HEADER][0];
  const appEmail = e.namedValues[CORE_APP_EMAIL_APP_HEADER] ? e.namedValues[CORE_APP_EMAIL_APP_HEADER][0] : '';
  
  if (clinicName && clinicEmail) {
    updateCoreMailAddress(clinicName, clinicEmail, appEmail);
  } else {
    logMessage(`handleCoreAppFormSubmit WAARSCHUWING: Clinicnaam of clinic email ontbreken in formulierinzending.`);
  }
}

function processCoreAppManualEdit(e) {
  const sheet = e.source.getActiveSheet();
  const editedRow = e.range.getRow();
  if (editedRow <= 1) return; // Ignore header edits

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[String(header).trim()] = index;
  });

  const clinicName = rowData[headerMap[CORE_APP_CLINIC_HEADER]];
  const clinicEmail = rowData[headerMap[CORE_APP_EMAIL_CLINIC_HEADER]];
  const appEmail = rowData[headerMap[CORE_APP_EMAIL_APP_HEADER]];

  if (clinicName && clinicEmail) {
    updateCoreMailAddress(clinicName, clinicEmail, appEmail);
  } else {
    logMessage(`processCoreAppManualEdit WAARSCHUWING: Clinicnaam of clinic email ontbreken op handmatig bewerkte rij ${editedRow}.`);
  }
}

function updateCoreMailAddress(clinicName, clinicEmail, appEmail) {
  const finalAppEmail = (appEmail && appEmail.trim() !== '') ? appEmail.trim() : clinicEmail.trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let participantFoundAndUpdated = false;
  
  logMessage(`Bijwerken CORE-mailadres voor clinic '${clinicName}', clinic email '${clinicEmail}' naar '${finalAppEmail}'.`);

  for (const sheetName of[OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME]) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        logMessage(`WAARSCHUWING: Respons-sheet '${sheetName}' niet gevonden.`);
        continue;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const clinicColIdx = headers.indexOf(FORM_EVENT_QUESTION_TITLE);
    const emailColIdx = headers.indexOf(FORM_EMAIL_QUESTION_TITLE);
    const coreMailColIdx = headers.indexOf(FORM_CORE_MAIL_HEADER);

    if (clinicColIdx === -1 || emailColIdx === -1 || coreMailColIdx === -1) {
        logMessage(`WAARSCHUWING: Een van de verplichte kolommen (Clinic, Email, CORE-mailadres) ontbreekt in tabblad '${sheetName}'.`);
        continue;
    }

    for (let i = 0; i < data.length; i++) {
      const rowClinicName = String(data[i][clinicColIdx] || '').split(' (')[0].trim();
      const rowEmail = String(data[i][emailColIdx] || '').trim().toLowerCase();

      if (rowClinicName === clinicName && rowEmail === String(clinicEmail).trim().toLowerCase()) {
        const currentCoreMail = String(data[i][coreMailColIdx] || '').trim();
        if (currentCoreMail !== finalAppEmail) {
            sheet.getRange(i + 2, coreMailColIdx + 1).setValue(finalAppEmail); // +2 for 1-based index and header row
            logMessage(`CORE-mailadres voor ${clinicEmail} in clinic '${clinicName}' bijgewerkt naar '${finalAppEmail}' in tabblad '${sheetName}'.`);
        } else {
            logMessage(`CORE-mailadres voor ${clinicEmail} in clinic '${clinicName}' is al '${finalAppEmail}' in tabblad '${sheetName}'. Geen wijziging nodig.`);
        }
        participantFoundAndUpdated = true;
        break; // Participant found and (potentially) updated in this sheet
      }
    }
    if (participantFoundAndUpdated) break; // If updated in one sheet, no need to check other response sheets
  }

  if (!participantFoundAndUpdated) {
    logMessage(`FOUT: Deelnemer met email '${clinicEmail}' voor clinic '${clinicName}' niet gevonden in enige respons-sheet om CORE-mailadres bij te werken.`);
  }
}