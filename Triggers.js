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
    .addItem('Archiveer doorgestreepte deelnemers', 'archiveStrikethroughParticipants')
    .addSeparator()
    .addItem('Update pop-ups voor alle formulieren', 'updateAllFormDropdowns')
    .addItem('Herstel alle agenda-items', 'recreateAllCalendarEvents')
    .addItem('Check of alle permissies zijn toegekend', 'forceAuthorization')
    .addToUi();
  
  // Check if the correct Gmail alias is available (indicates correct account)
  try {
    const desiredAlias = EMAIL_SENDER_ALIAS;
    const availableAliases = GmailApp.getAliases();
    
    if (!availableAliases.includes(desiredAlias)) {
      const currentUser = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'onbekend';
      // Show a one-time warning toast (non-intrusive)
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `⚠️ Je bent ingelogd als ${currentUser}. Om mails correct te verzenden dien je ingelogd te zijn als joost@thermoclinics.nl. TIP: Gebruik een incognito venster!`,
        '⚠️ Verkeerd Google Account',
        10 // Show for 10 seconds
      );
    }
  } catch (e) {
    // GmailApp.getAliases() failed - this typically means Gmail permissions aren't granted yet
    // Only show authorization message if it's actually an authorization error
    if (e.message && (e.message.includes('authorize') || e.message.includes('permission') || e.message.includes('access'))) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Ga naar "Thermoclinics Tools" → "Check of alle permissies zijn toegekend" om alle functies te kunnen gebruiken.',
        'ℹ️ Autorisatie vereist',
        8
      );
    }
    // For other errors (like wrong account type), silently ignore - the mail merge dialogs will show their own warnings
  }
}

function masterOnEdit(e) {
  try {
    const sheetName = e.source.getActiveSheet().getName();
    Logger.log(`masterOnEdit triggered on sheet: ${sheetName}`);

    switch (sheetName) {
      case DATA_CLINICS_SHEET_NAME:
        Logger.log(`Routing to handleEventChange(), handleClinicTypeChange(), syncCalendarEventFromSheet(), updateAllFormDropdowns(), autoSortSheet(), and removeBlankRows().`);
        handleEventChange(e); // Handle date, time, and location changes with folder renames
        handleClinicTypeChange(e); // Handle clinic type changes (Open <-> Besloten)
        syncCalendarEventFromSheet(e.range.getRow()); // Sync calendar for any edit (including max seats changes)
        updateAllFormDropdowns();
        autoSortSheet(e.source.getActiveSheet());
        removeBlankRows(e.source.getActiveSheet());
        break;
      case CORE_APP_SHEET_NAME:
        Logger.log(`Routing to processCoreAppManualEdit().`);
        processCoreAppManualEdit(e);
        break;
      case OPEN_FORM_RESPONSE_SHEET_NAME:
      case BESLOTEN_FORM_RESPONSE_SHEET_NAME:
        Logger.log(`Routing to handleParticipantNameChange(), autoSortSheet(), and removeBlankRows().`);
        handleParticipantNameChange(e);
        autoSortSheet(e.source.getActiveSheet());
        removeBlankRows(e.source.getActiveSheet());
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

/**
 * Automatically sorts a sheet based on its type.
 * - Open/Besloten: sorts by column A (timestamp) descending (newest first)
 * - Data Clinics: sorts by column A then column B ascending
 * Only sorts the data range (excluding the header row).
 * @param {Sheet} sheet - The sheet to sort
 */
function autoSortSheet(sheet) {
  const sheetName = sheet.getName();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  // Only sort if there's more than just the header row
  if (lastRow > 1 && lastCol > 0) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    
    if (sheetName === DATA_CLINICS_SHEET_NAME) {
      // Data Clinics: sort by column A, then column B (ascending)
      dataRange.sort([
        { column: 1, ascending: true },
        { column: 2, ascending: true }
      ]);
      Logger.log(`Auto-sorted sheet "${sheetName}" by column A then B (ascending).`);
    } else {
      // Open/Besloten: sort by timestamp descending (newest first)
      dataRange.sort({ column: 1, ascending: false });
      Logger.log(`Auto-sorted sheet "${sheetName}" by timestamp (column A, newest first).`);
    }
  }
}

/**
 * Removes blank rows from a sheet based on key columns.
 * - Data Clinics: row is blank if columns A through E are all empty
 * - Open/Besloten: row is blank if column A (timestamp) is empty
 * Iterates from bottom to top to avoid index shifting issues.
 * Uses banded ranges to detect the actual table extent (including formatted empty rows).
 * @param {Sheet} sheet - The sheet to clean
 */
function removeBlankRows(sheet) {
  const sheetName = sheet.getName();
  
  // Try to find the table extent from banded ranges (table formatting)
  let lastRow = sheet.getLastRow();
  const bandings = sheet.getBandings();
  if (bandings && bandings.length > 0) {
    // Use the largest banded range to determine table extent
    for (const banding of bandings) {
      const bandingLastRow = banding.getRange().getLastRow();
      if (bandingLastRow > lastRow) {
        lastRow = bandingLastRow;
      }
    }
  }
  
  if (lastRow <= 1) return; // Nothing to clean
  
  let deletedCount = 0;
  
  // Iterate from bottom to top to avoid index issues when deleting
  for (let row = lastRow; row >= 2; row--) {
    let isBlank = false;
    
    if (sheetName === DATA_CLINICS_SHEET_NAME) {
      // Data Clinics: check columns A through E
      const values = sheet.getRange(row, 1, 1, 5).getValues()[0];
      isBlank = values.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
    } else {
      // Open/Besloten: check column A (timestamp)
      const colA = sheet.getRange(row, 1).getValue();
      isBlank = colA === null || colA === undefined || String(colA).trim() === '';
    }
    
    if (isBlank) {
      sheet.deleteRow(row);
      deletedCount++;
    }
  }
  
  if (deletedCount > 0) {
    Logger.log(`Removed ${deletedCount} blank row(s) from sheet "${sheetName}".`);
  }
}