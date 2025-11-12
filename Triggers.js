// This file contains the main trigger functions (onOpen, masterOnEdit, masterOnFormSubmit)
// which act as routers for the script's execution.

function onOpen() {
  SpreadsheetApp.getUi()
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
}

function masterOnEdit(e) {
  try {
    const sheetName = e.source.getActiveSheet().getName();
    Logger.log(`masterOnEdit triggered on sheet: ${sheetName}`);

    switch (sheetName) {
      case DATA_CLINICS_SHEET_NAME:
        Logger.log(`Routing to handleTimeChange(), updateAllFormDropdowns(), and syncCalendarEventFromSheet().`);
        handleTimeChange(e); // Handle time changes and folder renames first
        updateAllFormDropdowns();
        syncCalendarEventFromSheet(e.range.getRow());
        break;
      case CORE_APP_SHEET_NAME:
        Logger.log(`Routing to processCoreAppManualEdit().`);
        processCoreAppManualEdit(e);
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