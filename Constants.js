// SCRIPT ID: 13pP6YA7hDs54PWbUbbktLPn3B8wa6wvjSxya-DYkyUgnC6c1ptL8j6X9
// This script is bound to the Google Sheet where FORM RESPONSES are stored.

// --- GLOBAL CONFIGURATION (accessible by all functions in this script) ---
const OPEN_FORM_ID = '1bZay8Zqv2IVg-hZaJiD-Hg4GEFJ4MdXAqmgPG_UudRw';
const BESLOTEN_FORM_ID = '1WY5L0t7ArvVzS7iTV3TG_kXDirW9ZQcsNrcQSG97l9M';
const CORE_APP_FORM_ID = '1nPqe00RzYA99yNIDspxThsTty1HODmvYNLJmCzprshQ';

const DATA_CLINICS_SPREADSHEET_ID = '15JjYK3O4k9IKFNxocC43D1ZhiH5g0Fd36ULRCfkLNEk';
const DATA_CLINICS_SHEET_NAME = 'Data clinics';
const QUESTION_TITLE_TO_UPDATE = 'Datum, tijdstip en locatie waarop je wilt deelnemen';
const CORE_APP_QUESTION_TITLE = 'Voor welke clinic heb je je opgegeven?'; // Specific title for the 3rd form
const PARENT_EVENT_FOLDER_ID = '11cNspj1CPYmUN7MUunHgg6Pv2QyhCwIT';
const CONFIRMATION_EMAIL_TEMPLATE_ID = '1MaoY-S2FgsajUv6Vp9a_s8TXjql7lka95t9DNvQy21k'; // Legacy template - kept for backward compatibility
const OPEN_CONFIRMATION_EMAIL_TEMPLATE_ID = '1MaoY-S2FgsajUv6Vp9a_s8TXjql7lka95t9DNvQy21k'; // Template for Open clinic confirmations
const BESLOTEN_CONFIRMATION_EMAIL_TEMPLATE_ID = '1MaoY-S2FgsajUv6Vp9a_s8TXjql7lka95t9DNvQy21k'; // Template for Besloten clinic confirmations (TODO: Update with actual template ID)
const CORE_APP_REMINDER_TEMPLATE_ID = '10CVavxSJnjg72LjTtMKaRnI4xN_KIrbbK8mp_QX94PE'; // Added constant for reminder template
const FALLBACK_EMAIL_SENDER_NAME = 'Thermoclinics';
const TARGET_CALENDAR_ID = 'c37540c489700aad336b5d9b240759b5c8b26b4473d6f4641a938b38be6344b0@group.calendar.google.com';
const DEFAULT_EVENT_DURATION_HOURS = 3;

// --- CONFIGURATION FOR MAIL MERGE & RESPONSE SHEETS ---
const MAIL_TEMPLATE_FOLDER_ID = '1utGcVyAQDkcLgAp0YC_vM9bdAscFp7ZT';
const GENERIC_ATTACHMENTS_FOLDER_ID = '1DZ1yJQtLMTTBoS2_hNu89rhaUt6GuS2B';
const EXCEL_IMPORT_FOLDER_ID = '1iJg3L3zWL2Km7Tjxn1JKDqcPS9u3BEMF'; // Folder for Excel imports
const OPEN_FORM_RESPONSE_SHEET_NAME = 'Open Form Responses';
const BESLOTEN_FORM_RESPONSE_SHEET_NAME = 'Besloten Form Responses';
const DEELNEMERNUMMER_HEADER = 'Deelnemernummer';
const DRIVE_FOLDER_ID_HEADER = 'Participant Folder ID';
const LOG_DOCUMENT_ID = '1-SaDBBm1R6Ethjc4E183ileZq7Me0nYzHKFChHPdEG4';
const CALENDAR_EVENT_ID_HEADER = 'Calendar Event ID';
const EVENT_FOLDER_ID_HEADER = 'Event Folder ID'; // NEW CONSTANT
const ARCHIVE_SHEET_NAME = 'ARCHIEF oudere clinics';
const ARCHIVE_PARTICIPANTS_SHEET_NAME = 'ARCHIEF deelnemers';
const NON_PARTICIPANT_EMAILS_SHEET_NAME = 'Non-participant emails';

const CORE_APP_SHEET_NAME = 'CORE-app geïnstalleerd';
const CORE_APP_CLINIC_HEADER = 'Voor welke clinic heb je je opgegeven?';
const CORE_APP_EMAIL_CLINIC_HEADER = 'Met welk email-adres heb je je opgegeven voor deze clinic?';
const CORE_APP_EMAIL_APP_HEADER = 'Bij het activeren van de CORE app heb ik het volgende e-mailadres gebruikt';

const DATE_COLUMN_INDEX = 1;
const TIME_COLUMN_INDEX = 2;
const LOCATION_COLUMN_INDEX = 3;
const MAX_SEATS_COLUMN_INDEX = 4;
const BOOKED_SEATS_COLUMN_INDEX = 5;
const TYPE_COLUMN_INDEX = 6;
const DATA_CLINICS_START_ROW = 2;

const DATE_FORMAT_YYYYMMDD = 'yyyyMMdd';
const FORMATTING_TIME_ZONE = Session.getScriptTimeZone();

// Form question titles (MUST match exactly)
const FORM_EVENT_QUESTION_TITLE = QUESTION_TITLE_TO_UPDATE;
const FORM_FIRST_NAME_QUESTION_TITLE = 'Voornaam';
const FORM_LAST_NAME_QUESTION_TITLE = 'Achternaam';
const FORM_EMAIL_QUESTION_TITLE = 'Email';
const FORM_PHONE_QUESTION_TITLE = 'Telefoonnummer';
const FORM_DOB_QUESTION_TITLE = 'Geboortedatum';
const FORM_CITY_QUESTION_TITLE = 'Woonplaats';
const FORM_OPMERKING_QUESTION_TITLE = 'Opmerkingen';
const FORM_MOTIVATIE_QUESTION_TITLE = 'Wat is je motivatie om deel te nemen?';
const FORM_REG_METHOD_QUESTION_TITLE = 'Manier van Inschrijving';
const FORM_CORE_MAIL_HEADER = 'CORE-mailadres';
const DEFAULT_UNASSIGNED_PARTICIPANT_NAME = 'Geen persoon toegewezen';