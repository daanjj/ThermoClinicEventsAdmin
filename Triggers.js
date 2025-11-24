// This file contains the main trigger functions (onOpen, masterOnEdit, masterOnFormSubmit)
// which act as routers for the script's execution.

function onOpen() {
  const menu = SpreadsheetApp.getUi()
    .createMenu('Thermoclinics Tools')
    .addItem('Verstuur mail naar deelnemers', 'showMailMergeDialog')
    .addItem('Stuur reminder om CORE-app te installeren', 'showCoreReminderDialog')
    .addItem('Lees Excel-bestand in', 'showExcelImportDialog')
    .addItem('Maak deelnemerslijst', 'showParticipantListDialog') 
    .addSeparator()
    .addItem('Archiveer oudere clinics', 'runManualArchive')
    .addSeparator()
    .addItem('Update pop-ups voor alle formulieren', 'updateAllFormDropdowns')
    .addItem('Check of alle permissies zijn toegekend', 'forceAuthorization')
    .addToUi();
  
  // Check if user is logged in with the correct account and show warning if not
  try {
    const expectedEmail = "infothermoclinics@gmail.com";
    const activeUserEmail = Session.getEffectiveUser().getEmail();
    
    if (activeUserEmail !== expectedEmail) {
      // Show a one-time warning toast (non-intrusive)
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `⚠️ U bent ingelogd als ${activeUserEmail}. Voor mail merge moet u ingelogd zijn als ${expectedEmail}. TIP: Gebruik een incognito venster!`,
        '⚠️ Verkeerd Google Account',
        10 // Show for 10 seconds
      );
    }
  } catch (e) {
    // If we can't check the user (permission not granted), show a toast suggesting authorization
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Ga naar "Thermoclinics Tools" → "Check of alle permissies zijn toegekend" om alle functies te kunnen gebruiken.',
      'ℹ️ Autorisatie vereist',
      8
    );
  }
}

function masterOnEdit(e) {
  try {
    const sheetName = e.source.getActiveSheet().getName();
    Logger.log(`masterOnEdit triggered on sheet: ${sheetName}`);

    switch (sheetName) {
      case DATA_CLINICS_SHEET_NAME:
        Logger.log(`Routing to handleEventChange(), handleClinicTypeChange(), syncCalendarEventFromSheet(), and updateAllFormDropdowns().`);
        handleEventChange(e); // Handle date, time, and location changes with folder renames
        handleClinicTypeChange(e); // Handle clinic type changes (Open <-> Besloten)
        syncCalendarEventFromSheet(e.range.getRow()); // Sync calendar for any edit (including max seats changes)
        updateAllFormDropdowns();
        break;
      case CORE_APP_SHEET_NAME:
        Logger.log(`Routing to processCoreAppManualEdit().`);
        processCoreAppManualEdit(e);
        break;
      case OPEN_FORM_RESPONSE_SHEET_NAME:
      case BESLOTEN_FORM_RESPONSE_SHEET_NAME:
        Logger.log(`Routing to handleParticipantNameChange().`);
        handleParticipantNameChange(e);
        break;
    }
  } catch (err) {
    logMessage(`Fout in masterOnEdit: ${err.toString()}`);
    Logger.log(`Error in masterOnEdit: ${err.toString()}`);
  } finally {
    flushLogs(); // Ensures logs from any operation are saved
  }
}

function masterOnFormSubmit(e) {
  const sheetName = e.range.getSheet().getName();
  Logger.log(`masterOnFormSubmit triggered. Data landed on sheet: "${sheetName}"`);

  try {
    switch (sheetName) {
      case OPEN_FORM_RESPONSE_SHEET_NAME:
      case BESLOTEN_FORM_RESPONSE_SHEET_NAME:
        Logger.log(`Routing to processBooking().`);
        processBooking(e);
        break;
      case CORE_APP_SHEET_NAME:
        Logger.log(`Routing to handleCoreAppFormSubmit().`);
        handleCoreAppFormSubmit(e);
        break;
      default:
        Logger.log(`WARNING: Form submission on unhandled sheet "${sheetName}". No action taken.`);
        break;
    }
  } catch (err) {
    logMessage(`Fout in masterOnFormSubmit: ${err.toString()}`);
    Logger.log(`Error in masterOnFormSubmit: ${err.toString()}`);
  } finally {
    flushLogs(); // Ensures logs from any operation are saved
  }
}