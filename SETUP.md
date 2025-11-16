### **File Structure**

1. **Constants.js**  
   * Contains all const variables. This is your central configuration file.  
   * **Why:** A single, clean place for all configurable values. Easy to update.  
2. **Main.js (or Triggers.js)**  
   * onOpen()  
   * masterOnEdit(e)  
   * masterOnFormSubmit(e)  
   * **Why:** These are the primary entry points (triggers) for your script. They act as "routers" that delegate to other functions based on the event. They should ideally be lean and focused on routing.  
3. **Utils.js**  
   * logMessage(message)  
   * flushLogs()  
   * logToDocument(message)  
   * escapeRegExp(str)  
   * resolveTimeArithmeticPlaceholders(text, placeholderMap)  
   * getDutchDateString(dateObject)  
   * forceAuthorization() (This is a manual utility but fits well with general script health checks)  
   * **Why:** Generic helper functions that are used across various parts of your script. These are foundational.  
4. **EventsAndForms.js**  
   * handleTimeChange(e)  
   * syncCalendarEventFromSheet(rowNum)  
   * updateEventFolderIDs() (Another manual utility, but directly related to event folders)  
   * populateFormDropdown(formType)  
   * populateCoreAppFormDropdown()  
   * updateAllFormDropdowns()  
   * **Why:** Groups functions that manage clinic data (events, folders, calendar synchronization) and are responsible for populating the various form dropdowns based on the latest clinic data.  
5. **FormSubmission.js**  
   * processBooking(e)  
   * **Why:** This is the core logic for processing a new form response from the "Open" or "Besloten" forms. It updates seat counts, creates participant folders, sends confirmation emails, and triggers dropdown updates. It will call functions from EventsAndForms.js and MailMerge.js.  
6. **MailMerge.js**  
   * showMailMergeDialog()  
   * getAvailableClinicsList()  
   * getMailTemplates()  
   * getGenericAttachments()  
   * performMailMerge(selectedClinic, selectedTemplateId, selectedTemplateName, selectedAttachmentIds)  
   * mergeTemplateInDoc(templateId, placeholderMap)  
   * **Why:** All logic related to the Mail Merge feature, including the dialog, data retrieval for the dialog, and the actual merging and sending process. mergeTemplateInDoc is a key helper for this feature.  
7. **CoreApp.js**  
   * showCoreReminderDialog()  
   * getClinicsForReminder()  
   * sendCoreAppReminder(selectedClinic)  
   * handleCoreAppFormSubmit(e)  
   * processCoreAppManualEdit(e)  
   * updateCoreMailAddress(clinicName, clinicEmail, appEmail)  
   * **Why:** Encapsulates all functionality specific to the CORE app integration, including reminders and updating CORE app email addresses.  
8. **ExcelImport.js**  
   * getExcelFiles()  
   * showExcelImportDialog()  
   * processExcelFile(fileId)  
   * **Why:** A self-contained feature for importing participant data from Excel files.  
9. **Archiving.js**  
   * runManualArchive()  
   * runDailyArchive()  
   * archiveOldClinics(isManualTrigger)  
   * **Why:** Dedicated to the archiving process, which can be run manually or via a daily trigger. Now preserves participant data instead of deleting it.  
10. **VersionHistoryRecovery.js**  
    * recoverAllParticipantsToCSV()  
    * recoverParticipantsStepwise()  
    * showVersionHistoryInstructions()  
    * Various batch processing and recovery functions  
    * **Why:** Comprehensive system for recovering participant data from spreadsheet version history and archives. Essential for data recovery scenarios.  
11. **ParticipantLists.js**  
    * showParticipantListDialog()  
    * getClinicsForParticipantList()  
    * generateParticipantTable(selectedClinic)  
    * **Why:** Functions for generating participant reports or lists.

    