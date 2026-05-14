# Geautomatiseerd Systeem voor Clinic- & Evenementbeheer

Dit repository bevat een uitgebreide Google Apps Script-oplossing die is ontworpen om de registratie, het beheer en de communicatie rondom clinic-evenementen te automatiseren. Het maakt gebruik van Google Sheets, Google Forms, Google Calendar en Gmail om een gestroomlijnde workflow te bieden.

## Overzicht

Het systeem is gebouwd rond een centraal Google Sheet dat fungeert als de database voor alle clinic-evenementen. Google Formulieren worden gebruikt als publiek aanmeldformulier. Wanneer een nieuwe inschrijving binnenkomt, verzorgt een set geautomatiseerde scripts alles van bijwerken van plaatsaantallen en verzenden van bevestigingsmails tot het aanmaken van deelnemermapjes in Google Drive.

## Kernfuncties

* **Automatische verwerking van aanmeldingen**: Nieuwe aanmeldingen via Google Forms werken automatisch de beschikbaarheid van plaatsen bij in het centrale events-sheet.
* **Dynamische formulierupdates**: Keuzelijsten in Google Formulieren worden automatisch gevuld en bijgewerkt op basis van de events in het Google Sheet, zodat deelnemers alleen voor actuele evenementen kunnen inschrijven.
* **Synchronisatie met agenda**: Maakt en werkt Google Calendar-gebeurtenissen op basis van de gegevens in het sheet, inclusief deelnemersaantallen in de titel/omschrijving.
* **Gepersonaliseerde Mail Merge**: Een ingebouwde mailmerge-functie voor gepersonaliseerde bulkmails, met ondersteuning voor sjablonen, placeholders (zoals `<Voornaam>`) en bijlagen. Er zijn aparte bevestigingssjablonen voor Open en Besloten clinics.
* **CORE-app integratie**: Functionaliteit om communicatie te beheren voor deelnemers met de CORE Body Temperature Sensor app, inclusief herinneringen voor deelnemers die hun CORE-email nog niet hebben geregistreerd.
* **Deelnemerbeheer in Drive**:
  * Voor elk event wordt automatisch een hoofdmap aangemaakt in Drive en voor elke deelnemer een submap.
  * Importeert deelnemers uit Excel en maakt mappen en entries aan in de juiste response-sheet.
  * Genereert on-demand deelnemerslijsten in een dialoogvenster.
* **Geautomatiseerde archivering**: Een dagelijkse trigger archiveert oudere events en hun deelnemersgegevens, waarbij archiefgegevens behouden blijven (niet verwijderd).
* **Version history recovery**: Hulpmiddelen om deelnemersgegevens uit de versiegeschiedenis van het sheet te herstellen en te exporteren naar CSV.
* **Robuuste logging**: Belangrijke acties en fouten worden gelogd naar een centraal Google Document voor monitoring en debugging.

## Werking (kort)

1. **Data Hub**: Het `Data clinics` sheet is de single source of truth. Hier staan datum, tijd, locatie, capaciteit en type (Open of Besloten).
2. **Aanmelding**: Deelnemers melden zich aan via een van de Google Formulieren.
3. **Triggers**: `onFormSubmit` en `onEdit` triggers roepen routerfuncties aan die het werk verdelen over modules (verwerking, vellen bijwerken, agenda-sync, etc.).
4. **Verwerking**: Bij verwerking wordt:
   * het corresponderende clinic gevonden in het `Data clinics` sheet;
   * het aantal geboekte plaatsen verhoogd;
   * een eventmap en deelnemersmap gemaakt in Drive;
   * de deelnemer weggeschreven in de juiste response-sheet;
   * een bevestigingsmail gestuurd (Open vs Besloten);
   * de Google Calendar-event gesynchroniseerd;
   * dropdowns in de Google Formulieren bijgewerkt.

## Bestandsstructuur (kort)

* `Constants.js`: Centrale configuratie en IDs.
* `Triggers.js`: Hoofd entry points (`onOpen`, `masterOnEdit`, `masterOnFormSubmit`).
* `FormSubmission.js`: Logica voor verwerking van formulierinzendingen.
* `MailMerge.js`: Logica en UI voor mail merge.
* `COREApp.js`: CORE-app specifieke functies.
* `EventsAndForms.js`: Synchronisatie tussen Sheet, Calendar en formulieren.
* `ParticipantLists.js`: Genereren van deelnemerslijsten.
* `ExcelImport.js`: Importeren van deelnemers vanuit Excel.
* `Archiving.js`: Automatische archivering van oude clinics.
* `VersionHistoryRecovery.js`: Herstel uit versiegeschiedenis.
* `Utils.js`: Algemene helpers (logging, datumformattering, autorisatie).
* `*.html`: HTML-bestanden voor custom dialogen (Mail Merge, Participant Lists, etc.).

## Installatie & Configuratie

1. Kopieer alle `.js` en `.html` bestanden naar een nieuw Google Apps Script-project dat aan het betreffende spreadsheet is gebonden.
2. Configureer `Constants.js`: Vul alle placeholder-IDs in (Sheets, Forms, Drive folders, templates).
3. Schakel Advanced Services in: Zorg dat de Google Drive API is ingeschakeld in de Apps Script projectinstellingen.
4. OAuth-scopes: `appsscript.json` bevat de benodigde scopes (inclusief `userinfo.email`).
5. Machtigingen verlenen:
   * Run de `forceAuthorization` functie uit `Utils.js` in de editor om alle benodigde machtigingen te triggeren.
   * Mogelijk moet je de functie een tweede keer uitvoeren nadat je permissies hebt geaccepteerd.
6. Triggers instellen:
   * `masterOnFormSubmit` → trigger: On form submit
   * `masterOnEdit` → trigger: On edit
   * `runDailyArchive` → tijdgestuurde trigger (dagelijks, bijv. 02:00 - 03:00)

Na installatie verschijnt een menu `Thermoclinics Tools` in het sheet met toegang tot Mail Merge, Deelnemerslijsten, Archivering en overige handmatige functies.

## Opmerkingen over agenda-functies

* **Controle en synchronisatie**: Vanuit het menu is er nu `Controleer en synchroniseer agenda-items` — een veilige, niet-destructieve routine die ontbrekende agenda-items aanmaakt en afwijkende titels/locaties bijwerkt. De functie toont een samenvatting met het aantal aangemaakte en bijgewerkte items.
* **Herstel alle agenda-items (voorzichtig!)**: Er is tevens een functie `Herstel alle agenda-items` die bedoeld is voor volledige reset/migratie naar een nieuwe kalender. Vooraf moet je:
  1. Oude events handmatig uit de bron-kalender verwijderen.
  2. De kolom `Calendar Event ID` in het `Data clinics` sheet volledig leegmaken.
  3. Controleren dat `TARGET_CALENDAR_ID` naar de juiste doel-kalender verwijst.

  Deze functie zal vervolgens voor alle clinics nieuwe agenda-items aanmaken en de nieuwe event-IDs terugschrijven. Gebruik deze functie alleen als je zeker weet dat je alle bestaande agenda-items wilt vervangen.

---

*Dit project biedt een robuuste basis voor eventbeheer en is eenvoudig aan te passen en uit te breiden. Door configuratie centraal te houden en logica te splitsen in modules blijft het onderhoudbaar en schaalbaar.*
