# Changelog

## Recent Updates (November 2025)

### ✨ New Features

#### Clinic Type Change Handling
- **Automatic participant migration**: When changing a clinic type between Open ↔ Besloten, all participants are automatically moved to the correct response sheet
- **Data integrity**: Headers are validated before moving to ensure data consistency
- **Smart logging**: Detailed logs track participant movements and potential issues

**Impact**: Admins can now safely change clinic types without manual participant data migration.

#### Duplicate Form Submission Prevention
- **Smart duplicate detection**: System now detects if someone submits the same form multiple times for the same event
- **Automatic updates**: Duplicate submissions update the existing participant record instead of creating duplicates
- **Self-exclusion**: Current submission is excluded when checking for duplicates to avoid false positives
- **Folder preservation**: Participant folders and numbers are preserved when updating duplicates
- **Automatic cleanup**: Duplicate form submission rows are deleted after updating the original entry

**Impact**: Excel imports followed by form submissions no longer create duplicate entries. The system intelligently merges the data.

#### Additional Form Fields Support
- **Opmerkingen field**: Added support for comments/remarks field in forms
- **Motivatie field**: Added support for motivation field in forms
- **Placeholder mapping**: New placeholders `<Opmerking>` and `<Motivatie>` available in email templates
- **Update on duplicates**: Both fields are updated when duplicate submissions are detected

#### Participant Folder Auto-Renaming
- **Name change detection**: When participant names are edited (via Excel import update or form submission), their folder is automatically renamed
- **Format preservation**: Maintains the numbered format (e.g., "01 John Doe" → "01 Jane Doe")
- **Graceful error handling**: Logs warnings if folder rename fails without breaking the workflow

#### Duplicate Folder Detection
- **Warning system**: System now detects and warns when multiple folders exist with the same event name
- **First folder selection**: Automatically uses the first folder found and logs which one was selected
- **Consistency**: Applied across form submissions and Excel imports

#### Mail Merge Account Verification
- **Active user check**: Mail merge now verifies the active Google account is `infothermoclinics@gmail.com`
- **Warning dialog**: Shows clear warning if wrong account is detected with option to continue or cancel
- **Gmail alias verification**: Checks if `info@thermoclinics.nl` alias is available before sending
- **Graceful fallback**: Emails are sent without 'from' alias if not available, with detailed logging
- **Security**: Prevents accidental email sends from personal accounts

**Impact**: Ensures emails are sent from the correct account with the proper sender alias, improving brand consistency and preventing user errors.

### 🐛 Bug Fixes

#### Form Submission Folder Name Format
- **Fixed regex bug**: Corrected time formatting regex from `/:|\./g` to `/:|\.\|/g` to properly handle folder names
- **Consistency**: Applied same fix to both `FormSubmission.js` and `ExcelImport.js`

### 🔧 Technical Improvements

#### Constants Updates
- Added `FORM_OPMERKING_QUESTION_TITLE` for comments field
- Added `FORM_MOTIVATIE_QUESTION_TITLE` for motivation field

#### Triggers Enhancement
- `masterOnEdit()` now routes to `handleClinicTypeChange()` for type column changes in Data clinics sheet

#### OAuth Scopes
- **Added `userinfo.email` scope**: Required for `Session.getEffectiveUser()` to verify active user account
- **Updated `appsscript.json`**: Includes new scope for mail merge account verification
- **Enhanced `forceAuthorization()`**: Now checks userinfo.email permission and logs active user for debugging

## Updates (November 2024)

### ✨ New Features

#### Separate Confirmation Emails for Open vs Besloten Clinics
- **Added separate email templates** for Open and Besloten clinic confirmations
- **Smart template selection**: System automatically chooses the correct template based on clinic type
- **New constants**: `OPEN_CONFIRMATION_EMAIL_TEMPLATE_ID` and `BESLOTEN_CONFIRMATION_EMAIL_TEMPLATE_ID`
- **Enhanced logging**: Confirmation emails now indicate which template type was used

**Impact**: Participants now receive customized confirmation emails based on whether they registered for an Open or Besloten clinic.

#### Version History Recovery System
- **Complete participant recovery**: New `VersionHistoryRecovery.js` module with comprehensive data recovery capabilities
- **Stepwise processing**: Batch processing to handle large version histories without timeouts
- **CSV export**: Exports all historical participant data with timestamps and version information
- **Manual recovery instructions**: Built-in guide for manual version history access

**Functions added**:
- `recoverAllParticipantsToCSV()` - Quick export of current + archived participants
- `recoverParticipantsStepwise()` - Full version history recovery in batches
- `showVersionHistoryInstructions()` - User-friendly recovery guide

#### Enhanced Archive System
- **Participant preservation**: Archived participants are now preserved in `ARCHIVE_PARTICIPANTS_SHEET_NAME`
- **Strike-through formatting**: Original entries are marked with strike-through instead of deletion
- **Archive recovery**: Full integration with the version history recovery system

### 🔧 Technical Improvements

#### Form Submission Processing
- **Clinic type detection**: `processBooking()` now determines clinic type from the Data clinics sheet
- **Template routing**: Automatic selection of confirmation email template based on clinic type
- **Fallback handling**: Graceful fallback to original template if clinic type is unrecognized

#### Constants Organization
- **New template constants**: Separate constants for Open and Besloten confirmation templates
- **Archive constants**: New constants for participant archiving functionality
- **Backward compatibility**: Original `CONFIRMATION_EMAIL_TEMPLATE_ID` maintained for fallback

### 📋 Configuration Updates Required

To complete the separate email template setup:

1. **Create Besloten template**: Create a separate Google Doc template for Besloten clinic confirmations
2. **Update constant**: Replace the placeholder in `BESLOTEN_CONFIRMATION_EMAIL_TEMPLATE_ID` with the actual template ID
3. **Customize content**: Tailor each template's content to be specific to Open vs Besloten clinics

### 🚀 Next Steps

- [ ] Create separate Besloten confirmation email template
- [ ] Update `BESLOTEN_CONFIRMATION_EMAIL_TEMPLATE_ID` with actual template ID
- [ ] Test both confirmation email flows
- [ ] Document the version history recovery procedures for end users

## Previous Updates

_Earlier changes were implemented before this changelog was created._