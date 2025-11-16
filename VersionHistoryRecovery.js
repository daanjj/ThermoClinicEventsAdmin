// This file contains functions for recovering participant data from version history and archives.

/**
 * Recovers all participants ever listed in the Open and Besloten response sheets,
 * including current participants and archived participants.
 * Exports the data to a CSV file in Google Drive.
 * @returns {string} Success message with CSV file information
 */
function recoverAllParticipantsToCSV() {
  const logPrefix = "Participant History Recovery";
  logMessage(`----- START ${logPrefix} -----`);
  
  try {
    const allParticipants = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Header for the consolidated CSV
    const csvHeaders = [
      'Source', 'Timestamp', 'Email', 'CORE Email', 'First Name', 'Last Name', 
      'Event Name', 'Phone', 'Date of Birth', 'City', 'Participant Number', 
      'Participant Folder ID', 'Additional Info'
    ];
    
    logMessage("Collecting current participants from active response sheets...");
    
    // Collect from current response sheets
    [OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME].forEach(sheetName => {
      try {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          logMessage(`WARNING: Sheet '${sheetName}' not found.`);
          return;
        }
        
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
          logMessage(`INFO: Sheet '${sheetName}' is empty or has only headers.`);
          return; // Only headers or empty
        }
        
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const headerMap = createHeaderMap(headers);
        
        logMessage(`Processing ${data.length - 1} participants from '${sheetName}'`);
        
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const participant = extractParticipantData(row, headerMap, `Current - ${sheetName}`, headers);
          if (participant.email) { // Only add if has email
            allParticipants.push(participant);
          }
        }
      } catch (e) {
        logMessage(`ERROR processing sheet '${sheetName}': ${e.message}`);
      }
    });
    
    logMessage("Collecting archived participants...");
    
    // Collect from archived participants sheet
    try {
      const archiveSheet = ss.getSheetByName(ARCHIVE_PARTICIPANTS_SHEET_NAME);
      if (archiveSheet && archiveSheet.getLastRow() > 1) {
        const archiveData = archiveSheet.getDataRange().getValues();
        const archiveHeaders = archiveData[0];
        
        logMessage(`Processing ${archiveData.length - 1} archived participants`);
        
        for (let i = 1; i < archiveData.length; i++) {
          const row = archiveData[i];
          const sourceSheet = row[0] || 'Unknown'; // First column is 'Bron Sheet'
          const actualData = row.slice(1); // Remove the 'Bron Sheet' column
          const originalHeaders = archiveHeaders.slice(1); // Remove 'Bron Sheet' header
          const headerMap = createHeaderMap(originalHeaders);
          
          const participant = extractParticipantData(actualData, headerMap, `Archived - ${sourceSheet}`, originalHeaders);
          if (participant.email) { // Only add if has email
            allParticipants.push(participant);
          }
        }
      } else {
        logMessage("No archived participants found.");
      }
    } catch (e) {
      logMessage(`ERROR processing archived participants: ${e.message}`);
    }
    
    // Remove duplicates based on email + event combination
    const uniqueParticipants = removeDuplicateParticipants(allParticipants);
    
    logMessage(`Total unique participants found: ${uniqueParticipants.length}`);
    
    // Create CSV content
    const csvContent = createCSVContent(csvHeaders, uniqueParticipants);
    
    // Save to Google Drive
    const fileName = `All_Participants_History_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.csv`;
    const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
    
    // Save to the Excel import folder (or root if folder not accessible)
    let file;
    try {
      const folder = DriveApp.getFolderById(EXCEL_IMPORT_FOLDER_ID);
      file = folder.createFile(blob);
    } catch (e) {
      logMessage(`Could not save to designated folder, saving to Drive root: ${e.message}`);
      file = DriveApp.createFile(blob);
    }
    
    const fileUrl = file.getUrl();
    const successMessage = `SUCCESS: Exported ${uniqueParticipants.length} unique participants to CSV file: ${fileName}`;
    
    logMessage(successMessage);
    logMessage(`File URL: ${fileUrl}`);
    logMessage(`File ID: ${file.getId()}`);
    
    // Show result to user
    const ui = SpreadsheetApp.getUi();
    ui.alert('Export Complete', 
      `${successMessage}\n\nFile saved as: ${fileName}\n\nYou can find it in Google Drive. The file URL has been logged.`, 
      ui.ButtonSet.OK);
    
    return successMessage;
    
  } catch (e) {
    const errorMessage = `ERROR in ${logPrefix}: ${e.toString()}`;
    Logger.log(`${errorMessage}\n${e.stack}`);
    logMessage(errorMessage);
    SpreadsheetApp.getUi().alert('Error', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    throw e;
  } finally {
    logMessage(`----- END ${logPrefix} -----`);
    flushLogs();
  }
}

/**
 * Creates a header map for quick column lookup
 * @param {Array} headers Array of header strings
 * @returns {Object} Map of header name to column index
 */
function createHeaderMap(headers) {
  const map = {};
  headers.forEach((header, index) => {
    if (header) {
      map[String(header).trim()] = index;
    }
  });
  return map;
}

/**
 * Extracts participant data from a row using header mapping
 * @param {Array} row Data row
 * @param {Object} headerMap Header to index mapping
 * @param {string} source Source identifier
 * @param {Array} originalHeaders Original headers for additional info
 * @returns {Object} Participant data object
 */
function extractParticipantData(row, headerMap, source, originalHeaders) {
  // Standard form question titles from Constants.js
  const timestampCol = headerMap['Timestamp'] !== undefined ? headerMap['Timestamp'] : 
                       headerMap['Tijdstempel'] !== undefined ? headerMap['Tijdstempel'] : 0;
  
  const emailCol = headerMap[FORM_EMAIL_QUESTION_TITLE] !== undefined ? headerMap[FORM_EMAIL_QUESTION_TITLE] : -1;
  const coreEmailCol = headerMap[FORM_CORE_MAIL_HEADER] !== undefined ? headerMap[FORM_CORE_MAIL_HEADER] : -1;
  const firstNameCol = headerMap[FORM_FIRST_NAME_QUESTION_TITLE] !== undefined ? headerMap[FORM_FIRST_NAME_QUESTION_TITLE] : -1;
  const lastNameCol = headerMap[FORM_LAST_NAME_QUESTION_TITLE] !== undefined ? headerMap[FORM_LAST_NAME_QUESTION_TITLE] : -1;
  const eventCol = headerMap[FORM_EVENT_QUESTION_TITLE] !== undefined ? headerMap[FORM_EVENT_QUESTION_TITLE] : -1;
  const phoneCol = headerMap[FORM_PHONE_QUESTION_TITLE] !== undefined ? headerMap[FORM_PHONE_QUESTION_TITLE] : -1;
  const dobCol = headerMap[FORM_DOB_QUESTION_TITLE] !== undefined ? headerMap[FORM_DOB_QUESTION_TITLE] : -1;
  const cityCol = headerMap[FORM_CITY_QUESTION_TITLE] !== undefined ? headerMap[FORM_CITY_QUESTION_TITLE] : -1;
  const participantNumCol = headerMap[DEELNEMERNUMMER_HEADER] !== undefined ? headerMap[DEELNEMERNUMMER_HEADER] : -1;
  const folderIdCol = headerMap[DRIVE_FOLDER_ID_HEADER] !== undefined ? headerMap[DRIVE_FOLDER_ID_HEADER] : -1;
  
  // Extract data with safe defaults
  const participant = {
    source: source,
    timestamp: timestampCol >= 0 && row[timestampCol] ? formatTimestamp(row[timestampCol]) : '',
    email: emailCol >= 0 ? String(row[emailCol] || '').trim() : '',
    coreEmail: coreEmailCol >= 0 ? String(row[coreEmailCol] || '').trim() : '',
    firstName: firstNameCol >= 0 ? String(row[firstNameCol] || '').trim() : '',
    lastName: lastNameCol >= 0 ? String(row[lastNameCol] || '').trim() : '',
    eventName: eventCol >= 0 ? String(row[eventCol] || '').trim() : '',
    phone: phoneCol >= 0 ? String(row[phoneCol] || '').trim() : '',
    dateOfBirth: dobCol >= 0 && row[dobCol] ? formatDate(row[dobCol]) : '',
    city: cityCol >= 0 ? String(row[cityCol] || '').trim() : '',
    participantNumber: participantNumCol >= 0 ? String(row[participantNumCol] || '').trim() : '',
    participantFolderId: folderIdCol >= 0 ? String(row[folderIdCol] || '').trim() : '',
    additionalInfo: ''
  };
  
  // Collect any additional non-standard columns as extra info
  const standardColumns = new Set([
    timestampCol, emailCol, coreEmailCol, firstNameCol, lastNameCol, 
    eventCol, phoneCol, dobCol, cityCol, participantNumCol, folderIdCol
  ]);
  
  const additionalInfo = [];
  originalHeaders.forEach((header, index) => {
    if (!standardColumns.has(index) && index < row.length && row[index]) {
      additionalInfo.push(`${header}: ${String(row[index]).trim()}`);
    }
  });
  participant.additionalInfo = additionalInfo.join(' | ');
  
  return participant;
}

/**
 * Removes duplicate participants based on email + event combination
 * @param {Array} participants Array of participant objects
 * @returns {Array} Array of unique participants
 */
function removeDuplicateParticipants(participants) {
  const seen = new Set();
  const unique = [];
  
  participants.forEach(participant => {
    const key = `${participant.email.toLowerCase()}|${participant.eventName}`;
    if (!seen.has(key)) {
      seen.add(key);
      unique.push(participant);
    }
  });
  
  return unique;
}

/**
 * Creates CSV content from headers and participant data
 * @param {Array} headers CSV headers
 * @param {Array} participants Participant data objects
 * @returns {string} CSV content
 */
function createCSVContent(headers, participants) {
  const rows = [headers];
  
  participants.forEach(participant => {
    rows.push([
      participant.source,
      participant.timestamp,
      participant.email,
      participant.coreEmail,
      participant.firstName,
      participant.lastName,
      participant.eventName,
      participant.phone,
      participant.dateOfBirth,
      participant.city,
      participant.participantNumber,
      participant.participantFolderId,
      participant.additionalInfo
    ]);
  });
  
  // Convert to CSV format
  return rows.map(row => 
    row.map(cell => {
      const cellStr = String(cell || '');
      // Escape quotes and wrap in quotes if contains comma, quote, or newline
      if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
        return '"' + cellStr.replace(/"/g, '""') + '"';
      }
      return cellStr;
    }).join(',')
  ).join('\n');
}

/**
 * Formats a timestamp for CSV export
 * @param {*} timestamp Timestamp value
 * @returns {string} Formatted timestamp
 */
function formatTimestamp(timestamp) {
  try {
    if (timestamp instanceof Date) {
      return Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    }
    return String(timestamp);
  } catch (e) {
    return String(timestamp);
  }
}

/**
 * Formats a date for CSV export
 * @param {*} date Date value
 * @returns {string} Formatted date
 */
function formatDate(date) {
  try {
    if (date instanceof Date) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    return String(date);
  } catch (e) {
    return String(date);
  }
}

/**
 * Creates instructions for manual version history recovery
 * This function provides step-by-step instructions for accessing Google Sheets version history
 * since Google Apps Script cannot directly access revision history.
 */
function showVersionHistoryInstructions() {
  const instructions = `
MANUAL VERSION HISTORY RECOVERY INSTRUCTIONS

Unfortunately, Google Apps Script cannot directly access Google Sheets version history.
However, you can manually recover deleted participants using these steps:

1. OPEN THE MAIN SPREADSHEET:
   - Go to your main Thermoclinics spreadsheet in Google Sheets
   - Look for the 'Open Form Responses' and 'Besloten Form Responses' tabs

2. ACCESS VERSION HISTORY:
   - Click on File > Version history > See version history
   - Or use Ctrl+Alt+Shift+H (Windows) or Cmd+Option+Shift+H (Mac)

3. BROWSE HISTORICAL VERSIONS:
   - You'll see a timeline of all changes to the spreadsheet
   - Click on any version to view the spreadsheet as it was at that time
   - Look for versions from before participants were deleted

4. IDENTIFY DELETED PARTICIPANTS:
   - Compare older versions with the current version
   - Look for participants in the 'Open Form Responses' and 'Besloten Form Responses' tabs
   - Note down participant details from older versions

5. RECOVER DATA:
   - You can copy participant rows from old versions
   - Paste them back into the current spreadsheet
   - Or use the 'Restore this version' option if you want to revert entirely

6. ALTERNATIVE - USE THE ARCHIVE:
   - Check the '${ARCHIVE_PARTICIPANTS_SHEET_NAME}' tab in your spreadsheet
   - Many deleted participants may already be preserved there
   - Run the 'Export All Participants' function to get a complete CSV

For automated recovery, run the 'Export All Participants to CSV' function from the menu.
This will collect all current and archived participants into a single CSV file.
  `;
  
  SpreadsheetApp.getUi().alert('Version History Recovery Instructions', instructions, SpreadsheetApp.getUi().ButtonSet.OK);
  logMessage("Version history instructions displayed to user");
}

/**
 * Counts the number of versions available in the current spreadsheet
 * This helps determine the scope of the version history recovery operation
 */
function countSpreadsheetVersions() {
  const logPrefix = "Version Counter";
  logMessage(`----- START ${logPrefix} -----`);
  
  try {
    // Get the current spreadsheet file ID
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const fileId = spreadsheet.getId();
    
    logMessage(`Checking version history for spreadsheet: ${spreadsheet.getName()}`);
    logMessage(`File ID: ${fileId}`);
    
    // Use Drive API to list all revisions
    const revisions = Drive.Revisions.list(fileId);
    const versionCount = revisions.items.length;
    
    // Get some details about the versions
    const oldestRevision = revisions.items[0];
    const newestRevision = revisions.items[revisions.items.length - 1];
    
    const oldestDate = new Date(oldestRevision.modifiedDate);
    const newestDate = new Date(newestRevision.modifiedDate);
    
    const summary = `
VERSION HISTORY SUMMARY

Total Versions Found: ${versionCount}

Oldest Version: ${Utilities.formatDate(oldestDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')}
Newest Version: ${Utilities.formatDate(newestDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')}

Time Span: ${Math.round((newestDate - oldestDate) / (1000 * 60 * 60 * 24))} days

This means the version history recovery would process ${versionCount} versions.
Based on the small size and stable structure you mentioned, this should be very manageable.

Estimated processing time: ${Math.ceil(versionCount / 10)} - ${Math.ceil(versionCount / 5)} minutes
    `;
    
    // Log the summary and also console.log for Apps Script editor
    logMessage(summary);
    console.log(summary);
    
    // Try to show UI alert, but don't fail if not available
    try {
      SpreadsheetApp.getUi().alert('Version History Count', summary, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (uiError) {
      console.log("UI not available - results logged above");
    }
    
    return versionCount;
    
  } catch (e) {
    const errorMessage = `ERROR in ${logPrefix}: ${e.toString()}`;
    Logger.log(`${errorMessage}\n${e.stack}`);
    logMessage(errorMessage);
    console.log(errorMessage);
    
    // Try to show UI alert, but don't fail if not available
    try {
      SpreadsheetApp.getUi().alert('Error', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (uiError) {
      console.log("UI not available - error logged above");
    }
    return 0;
  } finally {
    logMessage(`----- END ${logPrefix} -----`);
    flushLogs();
  }
}

/**
 * COMPLETE HISTORICAL RECOVERY - STEPWISE APPROACH
 * Processes versions in batches to avoid timeout. Call multiple times to process all versions.
 * @param {number} startVersion - Version to start from (1-based, default: 1)
 * @param {number} batchSize - Number of versions to process in this batch (default: 30)
 * @param {string} existingDataFileId - File ID of existing data to append to (optional)
 */
function recoverParticipantsStepwise(startVersion = 1, batchSize = 30, existingDataFileId = null) {
  const logPrefix = `Stepwise Recovery (${startVersion}-${Math.min(startVersion + batchSize - 1, 999)})`;
  logMessage(`----- START ${logPrefix} -----`);
  
  const startTime = new Date();
  let processedVersions = 0;
  let totalParticipants = 0;
  
  try {
    // Get the current spreadsheet file IDNow let me add t
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const fileId = spreadsheet.getId();
    const spreadsheetName = spreadsheet.getName();
    
    logMessage(`Starting stepwise recovery for: ${spreadsheetName}`);
    logMessage(`Batch: versions ${startVersion} to ${startVersion + batchSize - 1}`);
    console.log(`Processing batch: versions ${startVersion} to ${startVersion + batchSize - 1}`);
    
    // Get all revisions
    const revisions = Drive.Revisions.list(fileId);
    const totalVersionCount = revisions.items.length;
    const endVersion = Math.min(startVersion + batchSize - 1, totalVersionCount);
    
    logMessage(`Total versions available: ${totalVersionCount}`);
    logMessage(`Processing versions ${startVersion} to ${endVersion} (${endVersion - startVersion + 1} versions)`);
    console.log(`Total versions: ${totalVersionCount}, processing ${startVersion}-${endVersion}`);
    
    // Load existing data if resuming
    const allHistoricalParticipants = new Map(); // Key: email|eventName, Value: participant object with version info
    
    if (existingDataFileId) {
      try {
        logMessage(`Loading existing data from file ID: ${existingDataFileId}`);
        loadExistingParticipantData(existingDataFileId, allHistoricalParticipants);
        logMessage(`Loaded ${allHistoricalParticipants.size} existing participants`);
      } catch (loadError) {
        logMessage(`Warning: Could not load existing data: ${loadError.message}`);
      }
    }
    
    // Process specified range of versions (oldest first)
    for (let i = startVersion - 1; i < Math.min(startVersion + batchSize - 1, revisions.items.length); i++) {
      const revision = revisions.items[i];
      const versionDate = new Date(revision.modifiedDate);
      const versionNumber = i + 1;
      const batchPosition = versionNumber - startVersion + 1;
      
      processedVersions++;
      
      logMessage(`Processing version ${versionNumber} (batch ${batchPosition}/${batchSize}) from ${Utilities.formatDate(versionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')}`);
      console.log(`Processing version ${versionNumber} (batch ${batchPosition}/${batchSize})`);
      
      try {
        // Use exportLinks to download the revision as Excel format
        if (!revision.exportLinks || !revision.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']) {
          logMessage(`No Excel export link available for version ${versionNumber}, skipping`);
          continue;
        }
        
        const excelExportUrl = revision.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
        logMessage(`Downloading version ${versionNumber} from export URL`);
        
        // Download the revision using the export URL with rate limiting
        let response;
        let retryCount = 0;
        const maxRetries = 3;
        
        while (retryCount <= maxRetries) {
          try {
            response = UrlFetchApp.fetch(excelExportUrl, {
              headers: {
                'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
              }
            });
            
            if (response.getResponseCode() === 200) {
              break; // Success
            } else if (response.getResponseCode() === 429) {
              // Rate limited - wait and retry
              const waitTime = Math.min(10000 * Math.pow(2, retryCount), 30000); // Exponential backoff, max 30s
              logMessage(`Rate limited on version ${versionNumber}, waiting ${waitTime/1000}s before retry ${retryCount + 1}/${maxRetries}`);
              Utilities.sleep(waitTime);
              retryCount++;
            } else {
              throw new Error(`Failed to download revision: HTTP ${response.getResponseCode()}`);
            }
          } catch (fetchError) {
            if (fetchError.message.includes('429') && retryCount < maxRetries) {
              const waitTime = Math.min(10000 * Math.pow(2, retryCount), 30000);
              logMessage(`Rate limit error on version ${versionNumber}, waiting ${waitTime/1000}s before retry ${retryCount + 1}/${maxRetries}`);
              Utilities.sleep(waitTime);
              retryCount++;
            } else {
              throw fetchError;
            }
          }
        }
        
        if (response.getResponseCode() !== 200) {
          throw new Error(`Failed to download revision after ${maxRetries} retries: HTTP ${response.getResponseCode()}`);
        }
        
        const versionBlob = response.getBlob();
        
        // Add a small delay between requests to prevent rate limiting
        if (versionNumber < totalVersionCount) {
          Utilities.sleep(2000); // 2 second delay between versions
        }        // Create temporary Google Sheet from this version
        const tempFile = Drive.Files.insert({
          title: `[TEMP_HISTORY] Version_${versionNumber}_${Date.now()}`,
          mimeType: MimeType.GOOGLE_SHEETS,
          parents: [{id: DriveApp.getRootFolder().getId()}]
        }, versionBlob);
        
        const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
        
        // Process both response sheets from this version
        [OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME].forEach(sheetName => {
          try {
            const versionSheet = tempSpreadsheet.getSheetByName(sheetName);
            if (!versionSheet || versionSheet.getLastRow() < 2) return; // Skip if sheet doesn't exist or is empty
            
            const versionData = versionSheet.getDataRange().getValues();
            const headers = versionData[0];
            const headerMap = createHeaderMap(headers);
            
            // Process each participant in this version
            for (let rowIdx = 1; rowIdx < versionData.length; rowIdx++) {
              const row = versionData[rowIdx];
              const participant = extractParticipantData(row, headerMap, `Version ${versionNumber} - ${sheetName}`, headers);
              
              if (!participant.email) continue; // Skip rows without email
              
              const participantKey = `${participant.email.toLowerCase()}|${participant.eventName}`;
              
              // Check if we've seen this participant before
              if (!allHistoricalParticipants.has(participantKey)) {
                // First time seeing this participant
                participant.firstSeenVersion = versionNumber;
                participant.firstSeenDate = versionDate;
                participant.lastSeenVersion = versionNumber;
                participant.lastSeenDate = versionDate;
                participant.versionHistory = [versionNumber];
                participant.isCurrentlyActive = false; // Will be updated if found in latest version
                
                allHistoricalParticipants.set(participantKey, participant);
                totalParticipants++;
              } else {
                // Update last seen information
                const existingParticipant = allHistoricalParticipants.get(participantKey);
                existingParticipant.lastSeenVersion = versionNumber;
                existingParticipant.lastSeenDate = versionDate;
                existingParticipant.versionHistory.push(versionNumber);
                
                // Update any missing data with newer information
                if (!existingParticipant.firstName && participant.firstName) existingParticipant.firstName = participant.firstName;
                if (!existingParticipant.lastName && participant.lastName) existingParticipant.lastName = participant.lastName;
                if (!existingParticipant.phone && participant.phone) existingParticipant.phone = participant.phone;
                if (!existingParticipant.city && participant.city) existingParticipant.city = participant.city;
                if (!existingParticipant.coreEmail && participant.coreEmail) existingParticipant.coreEmail = participant.coreEmail;
              }
            }
          } catch (sheetError) {
            logMessage(`Warning: Could not process sheet '${sheetName}' in version ${versionNumber}: ${sheetError.message}`);
          }
        });
        
        // Clean up temporary file
        Drive.Files.remove(tempFile.id);
        
      } catch (versionError) {
        logMessage(`Error processing version ${versionNumber}: ${versionError.message}`);
        console.log(`Error processing version ${versionNumber}: ${versionError.message}`);
      }
      
      // Progress update every 5 versions
      if (batchPosition % 5 === 0 || versionNumber === endVersion) {
        const elapsed = (new Date() - startTime) / 1000;
        logMessage(`Batch progress: ${batchPosition}/${endVersion - startVersion + 1} versions in batch - ${Math.round(elapsed)}s elapsed`);
        console.log(`Batch progress: version ${versionNumber} - ${allHistoricalParticipants.size} total unique participants so far`);
      }
    }
    
    // Mark currently active participants by checking against current sheets
    markCurrentlyActiveParticipants(allHistoricalParticipants);
    
    // Convert to array and sort by first seen date
    const participantArray = Array.from(allHistoricalParticipants.values());
    participantArray.sort((a, b) => a.firstSeenDate - b.firstSeenDate);
    
    // Create enhanced CSV with historical information
    const csvContent = createHistoricalCSVContent(participantArray);
    
    // Save to Google Drive
    const fileName = `Complete_Participant_History_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.csv`;
    const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
    
    let file;
    try {
      const folder = DriveApp.getFolderById(EXCEL_IMPORT_FOLDER_ID);
      file = folder.createFile(blob);
    } catch (e) {
      logMessage(`Could not save to designated folder, saving to Drive root: ${e.message}`);
      file = DriveApp.createFile(blob);
    }
    
    const processingTime = Math.round((new Date() - startTime) / 1000);
    const currentCount = participantArray.filter(p => p.isCurrentlyActive).length;
    const deletedCount = participantArray.filter(p => !p.isCurrentlyActive).length;
    const nextStartVersion = endVersion + 1;
    const hasMoreVersions = endVersion < totalVersionCount;
    
    const successMessage = `
STEPWISE RECOVERY BATCH FINISHED

✅ Processed: versions ${startVersion} to ${endVersion} (${endVersion - startVersion + 1} versions)
✅ Total unique participants found: ${participantArray.length}
   - Currently active: ${currentCount}
   - Historically deleted: ${deletedCount}
✅ Processing time: ${processingTime} seconds
✅ CSV file created: ${fileName}
✅ File ID: ${file.getId()}
✅ File URL: ${file.getUrl()}

${hasMoreVersions ? 
  `⚠️ MORE VERSIONS TO PROCESS: ${totalVersionCount - endVersion} remaining
  
To continue, run: recoverParticipantsStepwise(${nextStartVersion}, ${batchSize}, "${file.getId()}")` :
  '✅ ALL VERSIONS PROCESSED! This is the complete participant history.'}
    `;
    
    logMessage(successMessage);
    console.log(successMessage);
    
    return successMessage;
    
  } catch (e) {
    const errorMessage = `ERROR in ${logPrefix}: ${e.toString()}`;
    Logger.log(`${errorMessage}\n${e.stack}`);
    logMessage(errorMessage);
    console.log(errorMessage);
    throw e;
  } finally {
    const totalTime = Math.round((new Date() - startTime) / 1000);
    logMessage(`----- END ${logPrefix} - Total time: ${totalTime} seconds -----`);
    console.log(`Recovery completed in ${totalTime} seconds`);
    flushLogs();
  }
}

/**
 * Loads existing participant data from a CSV file to resume processing
 * @param {string} fileId Google Drive file ID of existing CSV
 * @param {Map} participantMap Map to populate with existing data
 */
function loadExistingParticipantData(fileId, participantMap) {
  try {
    const file = DriveApp.getFileById(fileId);
    const csvContent = file.getBlob().getDataAsString();
    const lines = csvContent.split('\n');
    
    if (lines.length < 2) return; // No data rows
    
    const headers = lines[0].split(',');
    
    for (let i = 1; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;
      
      const values = parseCSVLine(line);
      if (values.length < headers.length) continue;
      
      // Extract participant data
      const email = values[0] || '';
      const eventName = values[1] || '';
      if (!email) continue;
      
      const participant = {
        email: email,
        eventName: eventName,
        firstName: values[2] || '',
        lastName: values[3] || '',
        coreEmail: values[4] || '',
        phone: values[5] || '',
        dateOfBirth: values[6] || '',
        city: values[7] || '',
        participantNumber: values[8] || '',
        participantFolderId: values[9] || '',
        isCurrentlyActive: values[10] === 'YES',
        firstSeenDate: values[11] ? new Date(values[11]) : new Date(),
        lastSeenDate: values[12] ? new Date(values[12]) : new Date(),
        firstSeenVersion: parseInt(values[13]) || 1,
        lastSeenVersion: parseInt(values[14]) || 1,
        versionHistory: values[15] ? values[15].toString().split(',').map(v => parseInt(v.trim())) : [1],
        source: values[16] || '',
        additionalInfo: values[17] || ''
      };
      
      const key = `${email.toLowerCase()}|${eventName}`;
      participantMap.set(key, participant);
    }
  } catch (e) {
    throw new Error(`Failed to load existing data: ${e.message}`);
  }
}

/**
 * Simple CSV line parser that handles quoted fields
 * @param {string} line CSV line to parse
 * @returns {Array} Array of field values
 */
function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++; // Skip next quote
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current);
  return result;
}

/**
 * Marks participants as currently active by checking current sheets
 * @param {Map} allHistoricalParticipants Map of all historical participants
 */
function markCurrentlyActiveParticipants(allHistoricalParticipants) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    [OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME].forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return;
      
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const headerMap = createHeaderMap(headers);
      
      const emailCol = headerMap[FORM_EMAIL_QUESTION_TITLE];
      const eventCol = headerMap[FORM_EVENT_QUESTION_TITLE];
      
      if (emailCol === undefined || eventCol === undefined) return;
      
      for (let i = 1; i < data.length; i++) {
        const email = String(data[i][emailCol] || '').trim().toLowerCase();
        const eventName = String(data[i][eventCol] || '').trim();
        
        if (email) {
          const key = `${email}|${eventName}`;
          if (allHistoricalParticipants.has(key)) {
            allHistoricalParticipants.get(key).isCurrentlyActive = true;
          }
        }
      }
    });
    
    logMessage("Marked currently active participants");
  } catch (e) {
    logMessage(`Warning: Could not mark current participants: ${e.message}`);
  }
}

/**
 * Creates CSV content with historical information
 * @param {Array} participants Array of participant objects with historical data
 * @returns {string} CSV content
 */
function createHistoricalCSVContent(participants) {
  const headers = [
    'Email', 'Event Name', 'First Name', 'Last Name', 'CORE Email', 'Phone', 
    'Date of Birth', 'City', 'Participant Number', 'Participant Folder ID',
    'Currently Active', 'First Seen Date', 'Last Seen Date', 'First Seen Version', 
    'Last Seen Version', 'Total Versions Seen', 'Source', 'Additional Info'
  ];
  
  const rows = [headers];
  
  participants.forEach(participant => {
    rows.push([
      participant.email,
      participant.eventName,
      participant.firstName,
      participant.lastName,
      participant.coreEmail,
      participant.phone,
      participant.dateOfBirth,
      participant.city,
      participant.participantNumber,
      participant.participantFolderId,
      participant.isCurrentlyActive ? 'YES' : 'NO',
      Utilities.formatDate(participant.firstSeenDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      Utilities.formatDate(participant.lastSeenDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      participant.firstSeenVersion,
      participant.lastSeenVersion,
      participant.versionHistory.length,
      participant.source,
      participant.additionalInfo
    ]);
  });
  
  // Convert to CSV format with proper escaping
  return rows.map(row => 
    row.map(cell => {
      const cellStr = String(cell || '');
      if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
        return '"' + cellStr.replace(/"/g, '""') + '"';
      }
      return cellStr;
    }).join(',')
  ).join('\n');
}

/**
 * CONVENIENCE FUNCTIONS FOR BATCH PROCESSING
 * These functions make it easy to process versions in quarters
 */

/**
 * Process first quarter of versions (1-34)
 */
function recoverParticipantsBatch1() {
  return recoverParticipantsStepwise(1, 34);
}

/**
 * Process second quarter of versions (35-68)
 * @param {string} existingDataFileId File ID from batch 1
 */
function recoverParticipantsBatch2(existingDataFileId) {
  return recoverParticipantsStepwise(35, 34, existingDataFileId);
}

/**
 * Process third quarter of versions (69-102)
 * @param {string} existingDataFileId File ID from batch 2
 */
function recoverParticipantsBatch3(existingDataFileId) {
  return recoverParticipantsStepwise(69, 34, existingDataFileId);
}

/**
 * Process fourth quarter of versions (103-137)
 * @param {string} existingDataFileId File ID from batch 3
 */
function recoverParticipantsBatch4(existingDataFileId) {
  return recoverParticipantsStepwise(103, 35, existingDataFileId);
}

/**
 * ORIGINAL FUNCTION RENAMED FOR BACKWARDS COMPATIBILITY
 */
function recoverAllParticipantsFromCompleteHistory() {
  console.log("⚠️ This function has been replaced with stepwise processing due to timeout issues.");
  console.log("Use recoverParticipantsBatch1() to start, then follow the instructions in the output.");
  return "Function replaced - use recoverParticipantsBatch1() to start stepwise processing";
}

/**
 * Test function to download and process a single revision using export links
 * This verifies the new approach works before running the full recovery
 */
function testSingleRevisionDownload() {
  const logPrefix = "Test Single Revision";
  logMessage(`----- START ${logPrefix} -----`);
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const fileId = spreadsheet.getId();
    
    const revisions = Drive.Revisions.list(fileId);
    if (revisions.items.length === 0) {
      throw new Error("No revisions found");
    }
    
    // Test with the first (oldest) revision
    const testRevision = revisions.items[0];
    const versionDate = new Date(testRevision.modifiedDate);
    
    logMessage(`Testing revision from ${Utilities.formatDate(versionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')}`);
    console.log(`Testing revision: ${testRevision.id}`);
    
    if (!testRevision.exportLinks || !testRevision.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']) {
      throw new Error("No Excel export link available for this revision");
    }
    
    const excelExportUrl = testRevision.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    logMessage(`Using export URL: ${excelExportUrl}`);
    
    // Download the revision
    const response = UrlFetchApp.fetch(excelExportUrl, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    
    const versionBlob = response.getBlob();
    const responseCode = response.getResponseCode();
    const contentLength = response.getHeaders()['Content-Length'] || 'unknown';
    
    logMessage(`✅ Successfully downloaded revision - Response code: ${responseCode}, Content-Length: ${contentLength}`);
    
    // Create temporary Google Sheet
    const tempFile = Drive.Files.insert({
      title: `[TEST_TEMP] Version_Test_${Date.now()}`,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{id: DriveApp.getRootFolder().getId()}]
    }, versionBlob);
    
    const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
    logMessage(`✅ Successfully created temporary sheet: ${tempFile.id}`);
    
    // Check if we can access the response sheets
    let foundSheets = [];
    [OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME].forEach(sheetName => {
      const sheet = tempSpreadsheet.getSheetByName(sheetName);
      if (sheet) {
        const rowCount = sheet.getLastRow();
        foundSheets.push(`${sheetName}: ${rowCount} rows`);
      }
    });
    
    // Clean up
    Drive.Files.remove(tempFile.id);
    
    const summary = `✅ TEST SUCCESSFUL!\nFound sheets: ${foundSheets.join(', ')}\nRevision download and processing works correctly.`;
    logMessage(summary);
    console.log(summary);
    
    return summary;
    
  } catch (e) {
    const errorMessage = `ERROR in ${logPrefix}: ${e.toString()}`;
    Logger.log(`${errorMessage}\n${e.stack}`);
    logMessage(errorMessage);
    console.log(errorMessage);
    return errorMessage;
  } finally {
    logMessage(`----- END ${logPrefix} -----`);
    flushLogs();
  }
}

/**
 * Debug function to examine revision structure and test API calls
 * This helps troubleshoot revision access issues before running the full recovery
 */
function debugRevisionStructure() {
  const logPrefix = "Debug Revisions";
  logMessage(`----- START ${logPrefix} -----`);
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const fileId = spreadsheet.getId();
    
    logMessage(`Debugging revisions for file ID: ${fileId}`);
    console.log(`Debugging revisions for file ID: ${fileId}`);
    
    const revisions = Drive.Revisions.list(fileId);
    console.log(`Found ${revisions.items.length} revisions`);
    
    // Examine first few revisions
    for (let i = 0; i < Math.min(3, revisions.items.length); i++) {
      const revision = revisions.items[i];
      console.log(`Revision ${i + 1}:`, JSON.stringify(revision, null, 2));
      
      // Test if we can access this revision
      try {
        logMessage(`Testing access to revision ${i + 1} with ID: ${revision.id}`);
        const versionBlob = Drive.Revisions.get(fileId, revision.id, {alt: 'media'});
        logMessage(`✅ Successfully accessed revision ${i + 1}`);
        console.log(`✅ Successfully accessed revision ${i + 1} - blob size: ${versionBlob.getSize ? versionBlob.getSize() : 'unknown'}`);
      } catch (accessError) {
        logMessage(`❌ Failed to access revision ${i + 1}: ${accessError.message}`);
        console.log(`❌ Failed to access revision ${i + 1}:`, accessError);
      }
    }
    
    const summary = `Examined ${Math.min(3, revisions.items.length)} of ${revisions.items.length} total revisions. Check console and logs for details.`;
    console.log(summary);
    logMessage(summary);
    
  } catch (e) {
    const errorMessage = `ERROR in ${logPrefix}: ${e.toString()}`;
    Logger.log(`${errorMessage}\n${e.stack}`);
    logMessage(errorMessage);
    console.log(errorMessage);
  } finally {
    logMessage(`----- END ${logPrefix} -----`);
    flushLogs();
  }
}

/**
 * Test function to verify the participant recovery works correctly
 * This function shows a summary of what would be exported without actually creating the CSV
 */
function testParticipantRecovery() {
  const logPrefix = "Test Participant Recovery";
  logMessage(`----- START ${logPrefix} -----`);
  
  try {
    let totalCount = 0;
    let summary = "PARTICIPANT RECOVERY TEST SUMMARY\n\n";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check current response sheets
    [OPEN_FORM_RESPONSE_SHEET_NAME, BESLOTEN_FORM_RESPONSE_SHEET_NAME].forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        const count = sheet.getLastRow() - 1; // Subtract header row
        totalCount += count;
        summary += `${sheetName}: ${count} participants\n`;
      } else {
        summary += `${sheetName}: Not found or empty\n`;
      }
    });
    
    // Check archived participants
    const archiveSheet = ss.getSheetByName(ARCHIVE_PARTICIPANTS_SHEET_NAME);
    if (archiveSheet && archiveSheet.getLastRow() > 1) {
      const count = archiveSheet.getLastRow() - 1; // Subtract header row
      totalCount += count;
      summary += `${ARCHIVE_PARTICIPANTS_SHEET_NAME}: ${count} archived participants\n`;
    } else {
      summary += `${ARCHIVE_PARTICIPANTS_SHEET_NAME}: Not found or empty\n`;
    }
    
    summary += `\nTOTAL PARTICIPANTS FOUND: ${totalCount}\n\n`;
    summary += "This is what would be exported to CSV if you run the full recovery function.";
    
    SpreadsheetApp.getUi().alert('Test Results', summary, SpreadsheetApp.getUi().ButtonSet.OK);
    logMessage(summary);
    
  } catch (e) {
    const errorMessage = `ERROR in ${logPrefix}: ${e.toString()}`;
    Logger.log(`${errorMessage}\n${e.stack}`);
    logMessage(errorMessage);
    SpreadsheetApp.getUi().alert('Test Error', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
  } finally {
    logMessage(`----- END ${logPrefix} -----`);
    flushLogs();
  }
}