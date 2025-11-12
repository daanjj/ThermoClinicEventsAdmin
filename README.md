# Automated Clinic & Event Management System

This repository contains a comprehensive Google Apps Script solution designed to automate the registration, management, and communication for clinic events. It leverages the power of Google Sheets, Google Forms, Google Calendar, and Gmail to create a seamless and efficient workflow.

## Overview

The system is built around a central Google Sheet that acts as a database for all clinic events. Google Forms are used as the public-facing interface for participants to register. When a new registration is submitted, a suite of automated scripts handles everything from updating seat counts and sending confirmation emails to creating dedicated participant folders in Google Drive.

## Core Features

*   **Automated Registration Processing**: New sign-ups via Google Forms automatically update available slots in the main events sheet.
*   **Dynamic Form Updates**: Google Form dropdowns for clinic selection are automatically populated and updated based on the events listed in the Google Sheet, ensuring participants can only sign up for active events.
*   **Calendar Synchronization**: Automatically creates and updates Google Calendar events based on the data in the sheet, including participant counts in the event description.
*   **Personalized Mail Merge**: A powerful, built-in mail merge feature allows for sending customized bulk emails to participants of a specific clinic. It supports HTML templates, placeholders (like `<Voornaam>`), and attachments.
*   **CORE App Integration**: Specific functionality to manage communication with participants regarding the CORE Body Temperature Sensor app, including sending reminders to those who haven't registered their app email.
*   **Participant Data Management**:
    *   Creates a unique Google Drive folder for each event and a subfolder for each participant upon registration.
    *   Imports participant data from Excel files, automatically creating folders and adding them to the appropriate response sheet.
    *   Generates on-demand participant lists with key information in a convenient dialog.
*   **Automated Archiving**: A daily, time-driven trigger automatically archives past events and their corresponding participant responses to keep the active sheets clean and performant.
*   **Robust Logging**: All major actions, errors, and communications are logged to a central Google Document for easy monitoring and debugging.

## How It Works

1.  **Data Hub**: The `Data clinics` Google Sheet is the single source of truth. Administrators add new clinics here, specifying date, time, location, capacity, and type ('Open' or 'Besloten').
2.  **Registration**: Participants register using one of the Google Forms (e.g., 'Open Form').
3.  **Trigger Execution**: An `onFormSubmit` trigger fires the `processBooking` function.
4.  **Processing**: The script:
    *   Finds the corresponding clinic in the `Data clinics` sheet.
    *   Increments the 'Booked Seats' count.
    *   Creates an event folder and a unique participant subfolder in Google Drive.
    *   Writes the new participant's data to the correct response sheet (e.g., `Open Form Responses`).
    *   Sends a personalized confirmation email using a template.
    *   Updates the Google Calendar event.
    *   Refreshes the dropdowns on all Google Forms to reflect the new seat availability.

## File Structure

The script is organized into logical modules for maintainability and clarity:

*   `Constants.js`: Central configuration file. All IDs for sheets, forms, folders, and templates are stored here.
*   `Main.js` (or `Triggers.js`): Contains the primary entry points for the script, such as `onOpen()`, `masterOnEdit(e)`, and `masterOnFormSubmit(e)`. These functions act as routers, delegating tasks to other modules.
*   `FormSubmission.js`: Handles the core logic for processing a new booking from a form submission.
*   `MailMerge.js`: Contains all logic for the mail merge feature, from displaying the user dialog to sending the emails.
*   `COREApp.js`: Encapsulates all functionality specific to the CORE app integration.
*   `EventsAndForms.js`: Manages the synchronization of data between the Google Sheet, Google Calendar, and Google Forms dropdowns.
*   `ParticipantLists.js`: Powers the feature for generating participant list reports.
*   `ExcelImport.js`: A self-contained feature for importing participant data from Excel files.
*   `Archiving.js`: Dedicated to the automated archiving process for old clinics.
*   `Utils.js`: Contains generic helper functions (like logging, date formatting, and authorization) used across the entire project.
*   `*.html`: HTML files that define the user interface for custom dialogs (e.g., Mail Merge, Participant Lists).

## Setup & Installation

1.  **Copy Files**: Copy all `.js` and `.html` files into a new Google Apps Script project bound to the Google Sheet that will receive the form responses.
2.  **Configure `Constants.js`**: Open `Constants.js` and replace all placeholder IDs with the actual IDs from your Google Drive folders, Google Sheets, Google Forms, and email templates. This is the most critical step.
3.  **Enable Advanced Services**: In the Apps Script Editor, go to `Services` and ensure the `Google Drive API` is enabled.
4.  **Grant Permissions**:
    *   Run the `forceAuthorization` function from the `Utils.js` file manually from the script editor.
    *   This will trigger a series of permission prompts from Google. You must approve all of them for the script to function correctly. You may need to run it a second time after granting initial permissions.
5.  **Set Up Triggers**: In the Apps Script Editor, go to the `Triggers` section (clock icon) and set up the following:
    *   `masterOnFormSubmit` to run from the spreadsheet **On form submit**.
    *   `masterOnEdit` to run from the spreadsheet **On edit**.
    *   `runDailyArchive` to run as a **Time-driven** trigger, daily, at a time of your choosing (e.g., 2-3 am).

Once set up, a new "ThermoClinic Menu" will appear in the Google Sheet UI, providing access to the manual features like Mail Merge, Participant Lists, and Archiving.

---

*This project provides a robust framework for event management that can be adapted for various use cases. By centralizing configuration and separating logic, it remains scalable and easy to maintain.*