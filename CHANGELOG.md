# Changelog

## Recent Updates (November 2024)

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