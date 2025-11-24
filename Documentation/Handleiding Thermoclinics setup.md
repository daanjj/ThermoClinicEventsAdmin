# Deelnemeradministratie Thermoclinics

1. # Inleiding

Dit document beschrijft de werking van het geautomatiseerde systeem voor de administratie en communicatie rondom de Thermoclinics. Het systeem is gebouwd in Google Apps Script en is gekoppeld aan het Google Sheet waarin de formulierreacties worden opgeslagen. Het doel is om aanmeldingen, communicatie, en archivering te stroomlijnen.

2. # Basisopzet

Het systeem is opgebouwd rond een aantal kerncomponenten in Google Sheets en Google Drive.

* **Data Clinics Spreadsheet:** Dit is het centrale "database" spreadsheet. Hierin staat een sheet genaamd Data clinics die fungeert als de enige bron van waarheid voor alle geplande clinics. Alle formulieren en processen halen hier hun informatie vandaan.  
* **Response Sheets:** In het hoofd-spreadsheet zijn er aparte tabbladen voor de reacties van de verschillende aanmeldformulieren (Open Form Responses, Besloten Form Responses, CORE-app geïnstalleerd).  
* **Google Drive Mappenstructuur:** Voor elke clinic wordt automatisch een map aangemaakt, en voor elke deelnemer binnen die clinic een submap. Dit zorgt voor een georganiseerde opslag van eventuele documenten per deelnemer.

3. # Het Menu "Thermoclinics Tools"

Bij het openen van het spreadsheet verschijnt er een extra menu genaamd "Thermoclinics Tools". Hieronder de functies:

* **Verstuur mail naar deelnemers:** Opent een dialoogvenster om een gepersonaliseerde e-mail (mail merge) te sturen naar alle deelnemers van een geselecteerde clinic.  
* **Stuur reminder om CORE-app te installeren:** Opent een venster om een herinneringsmail te sturen naar deelnemers van een specifieke clinic die nog niet hebben aangegeven de CORE-app te hebben geïnstalleerd.  
* **Lees Excel-bestand in voor besloten clinic:** Start de procedure om deelnemers voor een besloten clinic te importeren vanuit een Excel-bestand.  
* **Archiveer oudere clinics:** Start handmatig het archiveringsproces. Dit is dezelfde functie die ook elke nacht automatisch draait.  
* **Update pop-ups voor alle formulieren:** Werkt de keuzelijsten in alle gekoppelde Google Formulieren bij op basis van de actuele data in de Data clinics sheet. Dit gebeurt ook automatisch na elke aanmelding, maar kan handig zijn voor een handmatige synchronisatie.  
* **Check of alle permissies zijn toegekend:** Een functie voor de eerste ingebruikname of bij problemen. Deze functie triggert de autorisatievraag van Google voor alle benodigde services (Drive, Gmail, Calendar, User Info, etc.). Deze check omvat ook de verificatie van de userinfo.email permissie die nodig is voor account verificatie bij mail merge.

4. # Automatische Processen

Het systeem voert op de achtergrond een aantal taken volledig automatisch uit.

## 4.1	Aanmelding van een Deelnemer (via Google Form)

Wanneer een deelnemer zich via een van de Google Formulieren aanmeldt, gebeurt het volgende:

1. De gegevens komen binnen op de bijbehorende response sheet (Open of Besloten).  
2. Het systeem zoekt de geselecteerde clinic op in de Data clinics sheet.
3. **Controle op dubbele inschrijvingen:** Het systeem controleert of deze deelnemer (op basis van e-mailadres) al is ingeschreven voor dezelfde clinic:
   * **Bij dubbele inschrijving:** De bestaande deelnemersgegevens worden bijgewerkt met de nieuwe informatie uit het formulier. Het deelnemernummer en de map blijven behouden. De dubbele formulierinzending wordt verwijderd.
   * **Bij nieuwe inschrijving:** Het proces gaat verder zoals hieronder beschreven.
4. Het aantal geboekte plaatsen (Aantal boekingen) voor die clinic wordt met 1 verhoogd.  
5. Een uniek Deelnemernummer wordt gegenereerd op basis van het nieuwe aantal boekingen (bijv. 01, 02, 03).  
6. In Google Drive wordt een mappenstructuur aangemaakt of gebruikt:  
   * Een hoofdmap voor het evenement (bijv. 20250824 1000 Locatie X).  
   * Een submap voor de deelnemer (bijv. 03 Voornaam Achternaam).  
7. Het Deelnemernummer en de ID van de zojuist aangemaakte Drive-map worden teruggeschreven naar de response sheet in de rij van de deelnemer.  
8. Een automatische bevestigingsmail wordt verstuurd naar de deelnemer:
   * Voor **Open clinics** wordt de template gebruikt die is ingesteld bij `OPEN_CONFIRMATION_EMAIL_TEMPLATE_ID`
   * Voor **Besloten clinics** wordt de template gebruikt die is ingesteld bij `BESLOTEN_CONFIRMATION_EMAIL_TEMPLATE_ID`
9. De keuzelijsten in de aanmeldformulieren worden direct bijgewerkt om het nieuwe aantal beschikbare plaatsen te tonen.

## 4.2	Dagelijkse Archivering

Elke nacht om circa 03:00 uur 's nachts draait er een automatisch script dat het volgende doet:

1. Het controleert de Data clinics sheet op clinics waarvan de datum meer dan 30 dagen in het verleden ligt.  
2. Alle gegevens van deze oude clinics worden verplaatst van de Data clinics sheet naar de ARCHIEF oudere clinics sheet.  
3. **Deelnemersgegevens worden gepreserveerd:** Deelnemersgegevens worden NIET verwijderd maar gearchiveerd:
   - Alle deelnemers van gearchiveerde clinics worden gekopieerd naar een nieuwe sheet genaamd 'Archief deelnemers'
   - In de originele Open Form Responses en Besloten Form Responses sheets krijgen deze deelnemers doorgestreepte opmaak (strike-through)
   - De originele deelnemersgegevens blijven volledig intact voor referentie
   - In het archief wordt een extra kolom 'Bron Sheet' toegevoegd om bij te houden uit welk formulier de deelnemer oorspronkelijk kwam
4. **Agenda-items blijven behouden:** Kalendergebeurtenissen voor gearchiveerde clinics worden NIET verwijderd en blijven beschikbaar als historische documentatie.  
   Dit zorgt ervoor dat de operationele sheets schoon en relevant blijven, terwijl alle historische data behouden blijft.

### 4.2.1	Structuur van het Archief

**ARCHIEF oudere clinics sheet:**
- Bevat alle gearchiveerde clinic-gegevens met dezelfde kolomstructuur als de originele Data clinics sheet
- Behoudt alle informatie zoals datum, tijd, locatie, capaciteit, en kalender-IDs

**Archief deelnemers sheet:**
- Bevat alle gearchiveerde deelnemersgegevens 
- Heeft dezelfde kolomstructuur als de originele response sheets (Timestamp, Voornaam, Achternaam, Email, etc.)
- Bevat een extra kolom 'Bron Sheet' (als laatste kolom) die aangeeft uit welk formulier de deelnemer oorspronkelijk kwam ('Open Form Response' of 'Besloten Form Response')
- Alle timestamp-informatie en persoonlijke gegevens blijven volledig intact

**Doorgestreepte rijen in originele sheets:**
- Gearchiveerde deelnemers in de Open Form Responses en Besloten Form Responses sheets krijgen doorgestreepte opmaak
- Deze rijen worden NIET opnieuw gearchiveerd bij volgende archiveringsruns
- Data blijft beschikbaar voor referentie maar is visueel gemarkeerd als gearchiveerd

## 4.3	Synchronisatie met de Agenda

Elke wijziging in de Data clinics sheet (bijv. een nieuwe clinic toevoegen, of een aanmelding die het aantal deelnemers wijzigt) triggert een update naar de gedeelde Google Agenda (TARGET\_CALENDAR\_ID).

* Nieuwe rijen creëren een nieuw agenda-item.  
* Wijzigingen in tijd, locatie of aantal deelnemers werken het bestaande agenda-item bij. De titel van het item toont het aantal deelnemers, bijv. Thermoclinic op/bij Locatie X (5 deelnemers).

## 4.4	Automatische Synchronisatie bij Handmatige Wijzigingen

Het systeem detecteert automatisch handmatige wijzigingen in de spreadsheets en houdt alle gerelateerde elementen gesynchroniseerd:

### 4.4.1	Wijzigingen in Event-gegevens (Data clinics sheet)

Wanneer u handmatig de datum, tijd of locatie van een clinic wijzigt in de Data clinics sheet:

* **Agenda-item wordt bijgewerkt:** Het gekoppelde kalendergebeurtenis wordt automatisch aangepast met de nieuwe datum, tijd en locatie
* **Event-map wordt hernoemd:** De hoofdeventmap in Google Drive wordt automatisch hernoemd volgens het nieuwe formaat (bijv. van "20250824 1000 Locatie A" naar "20250825 1400 Locatie B")
* **Deelnemersgegevens worden gesynchroniseerd:** Alle deelnemers in de response sheets krijgen automatisch hun eventnaam bijgewerkt naar de nieuwe datum/tijd/locatie
* **Alle submappen blijven intact:** Deelnemersmappen binnen de eventmap behouden hun inhoud en structuur

### 4.4.1a Wijzigingen in Clinic Type (Open ↔ Besloten)

Wanneer u het type van een clinic wijzigt van Open naar Besloten (of vice versa) in de Data clinics sheet:

* **Automatische deelnemersmigratie:** Alle deelnemers van die clinic worden automatisch verplaatst van de ene response sheet (bijv. "Open Form Responses") naar de andere (bijv. "Besloten Form Responses")
* **Datavalidatie:** Het systeem controleert of de headers van beide sheets overeenkomen voordat de migratie plaatsvindt
* **Gedetailleerde logging:** Er wordt precies bijgehouden hoeveel deelnemers zijn verplaatst en van/naar welke sheet
* **Veiligheidscheck:** Als de headers niet overeenkomen, wordt de migratie afgebroken en krijgt u een waarschuwing

**Let op:** Deze functie is vooral handig als een clinic ten onrechte als Open of Besloten was aangemerkt en moet worden gecorrigeerd.

### 4.4.2	Wijzigingen in Deelnemersgegevens (Response sheets)

Wanneer u handmatig de voor- of achternaam van een deelnemer wijzigt in de Open Form Responses of Besloten Form Responses sheets:

* **Deelnemersmap wordt hernoemd:** De persoonlijke map van de deelnemer in Google Drive wordt automatisch hernoemd naar het nieuwe formaat (bijv. van "03 Jan Jansen" naar "03 Jane Janssen")
* **Mapinhoud blijft behouden:** Alle bestanden en documenten in de deelnemersmap blijven volledig intact
* **Deelnemernummer blijft gelijk:** Het unieke deelnemernummer verandert niet bij naamwijzigingen
* **Ook bij updates via Excel of formulier:** Als een bestaande deelnemer via Excel-import of een dubbele formulierinzending wordt bijgewerkt met een nieuwe naam, wordt de map automatisch hernoemd

### 4.4.3	Belangrijke Aandachtspunten

* **Alleen voor- en achternaam:** Automatische herbenaming van deelnemersmappen werkt alleen bij wijzigingen in de voor- en achternaam kolommen
* **Bestaande mappen:** Het systeem werkt alleen als de oorspronkelijke mappenstructuur intact is en correct benoemd volgens het verwachte formaat
* **Real-time updates:** Synchronisatie gebeurt direct bij het opslaan van wijzigingen in de spreadsheet
* **Dubbele mappen worden gedetecteerd:** Als er meerdere mappen met dezelfde naam bestaan, gebruikt het systeem de eerste gevonden map en logt een waarschuwing

## 4.5	Communicatie met Deelnemers (Mail Merge)

Via het menu kan een mail worden verstuurd naar de deelnemers van een specifieke clinic. Dit proces maakt gebruik van Google Docs als sjablonen.

### 4.5.1	Verificatie van het Juiste Google Account

**Belangrijk:** Mail merge moet altijd worden uitgevoerd vanuit het account **infothermoclinics@gmail.com** om te zorgen dat:
- Emails worden verzonden met de juiste afzender (info@thermoclinics.nl als alias)
- De branding consistent blijft
- Er geen emails vanuit een persoonlijk account worden verstuurd

**Automatische controle:**
- Bij het starten van een mail merge controleert het systeem automatisch of u bent ingelogd met het juiste account
- Als u bent ingelogd met een ander account, krijgt u een waarschuwingsvenster:
  - **"LET OP: mailmerge dient vanuit user infothermoclinics@gmail.com gedaan te worden en niet vanuit [uw email]. Weet u zeker dat u door wilt gaan? Mails zullen uit uw naam worden verzonden en niet vanuit info@thermoclinics.nl!"**
  - U kunt kiezen voor **"Ja"** (doorgaan op eigen risico) of **"Nee"** (annuleren en inloggen met het juiste account)
- Als u kiest voor "Nee", wordt de mail merge geannuleerd en kunt u opnieuw inloggen met het correcte account

**Gmail Alias Verificatie:**
- Het systeem controleert ook of de alias **info@thermoclinics.nl** beschikbaar is in het actieve Gmail account
- Als de alias niet beschikbaar is, worden emails verstuurd vanaf het standaard account (met een waarschuwing in de log)
- Voor optimale werking moet de alias correct zijn geconfigureerd in de Gmail-instellingen van infothermoclinics@gmail.com

### 4.5.2	Mailsjablonen

Alle mailsjablonen moeten als Google Document worden opgeslagen in de daarvoor bestemde Google Drive map (MAIL\_TEMPLATE\_FOLDER\_ID).

**Naamgeving:** Sjablonen die in de mail merge selectielijst verschijnen moeten het woord **"mailsjabloon"** in hun bestandsnaam bevatten (hoofdletterongevoelig).

### 4.5.3	Opbouw van een sjabloon

Een sjabloon heeft een vaste opbouw voor de eerste drie regels om de afzender en het onderwerp te definiëren:

1. **Regel 1:** Moet beginnen met Van:, gevolgd door de naam die als afzender van de e-mail moet verschijnen. Bijvoorbeeld: Van: Jouw Naam \- Thermoclinics.  
2. **Regel 2:** Moet beginnen met Onderwerp:, gevolgd door de onderwerpregel van de e-mail. Placeholders kunnen hierin gebruikt worden. Bijvoorbeeld: Onderwerp: Belangrijke informatie voor de clinic op \<Datum\>.  
3. **Regel 3:** Moet een **lege regel** zijn om de koptekst van de daadwerkelijke e-mailtekst te scheiden.  
4. **Vanaf Regel 4:** De inhoud van de e-mail zelf, waarin placeholders gebruikt kunnen worden.

Voorbeeld:

Generated code  
Van: Team Thermoclinics  
Onderwerp: Jouw deelname aan de clinic: \<Eventnaam\>

Beste \<Voornaam\>,

Hierbij de details voor de aankomende clinic...  
   

### 4.5.4	Naamgeving voor Sjablonen met Bijlagen

Als een mailsjabloon bedoeld is om **persoonlijke bijlagen** vanuit de deelnemersmap mee te sturen, moet de naam van het Google Doc-sjabloon het woord bijlage bevatten (hoofdletterongevoelig). Bijvoorbeeld: Mailsjabloon met persoonlijke bijlage.docx.

Wanneer een dergelijk sjabloon wordt gebruikt:

* Zoekt het systeem in de Drive-map van elke individuele deelnemer.  
* Alle bestanden die direct in die map staan, worden als bijlage aan de mail toegevoegd.  
* Nadat de mail succesvol is verstuurd, worden deze bestanden verplaatst naar een submap genaamd Reeds verstuurde bijlagen om te voorkomen dat ze opnieuw worden meegestuurd.

### 4.5.5	Generieke Bijlagen

Voor **alle** mailsjablonen (ook die zonder persoonlijke bijlagen) kunt u kiezen om generieke bijlagen toe te voegen:

* Bij het selecteren van een mailsjabloon in de mail merge dialoog verschijnt de vraag: **"Wil je een of meer generieke bestanden toevoegen?"**
* U kunt kiezen voor "Ja" of "Nee" (standaard is "Nee" geselecteerd)
* Bij "Ja" verschijnt een lijst met alle beschikbare bestanden uit de map **Algemene mailbijlagen** (GENERIC_ATTACHMENTS_FOLDER_ID)
* U kunt één of meer bestanden selecteren die aan alle emails worden toegevoegd
* Deze generieke bijlagen worden gecombineerd met eventuele persoonlijke bijlagen (bij sjablonen met "bijlage" in de naam)

**Voorbeeld gebruik:**
- Algemene informatieblaadjes over thermoclinics
- Routebeschrijvingen naar de locatie
- Algemene voorwaarden
- Veelgestelde vragen documenten

### 4.5.6	Beschikbare Placeholders

De volgende placeholders kunnen in de onderwerpregel en de body van het mailsjabloon worden gebruikt. Ze worden automatisch vervangen door de gegevens van de betreffende deelnemer.

**Deelnemersgegevens:**

* \<Voornaam\>  
* \<Achternaam\>  
* \<Email\>  
* \<Telefoonnummer\>  
* \<Geboortedatum\>  
* \<Woonplaats\>
* \<Opmerking\>: Opmerkingen die de deelnemer heeft ingevuld in het formulier
* \<Motivatie\>: Motivatie voor deelname (indien ingevuld in het formulier)
* \<Deelnemernummer\>  
* \<CORE-mailadres\>

**Eventgegevens:**

* \<Eventnaam\>: De volledige naam van de clinic, bijv. "zondag 25 augustus 2024 10:00-13:00, Amsterdam".  
* \<Datum\>: Alleen het datumgedeelte, bijv. "zondag 25 augustus 2024".  
* \<Tijd\>: De volledige tijdsaanduiding, bijv. "10:00-13:00".  
* \<Starttijd\>: Alleen de starttijd, bijv. "10:00".  
* \<Locatie\>: De locatie van de clinic, bijv. "Amsterdam".

### 4.5.5	Speciale Placeholders (Tijdrekenen)

Het is mogelijk om in de mailtekst te rekenen met de starttijd. Dit is handig voor het aangeven van bijvoorbeeld een inlooptijd. De syntax is: \<Starttijd \+/- AANTAL min\>.

Voorbeeld:

* Als \<Starttijd\> gelijk is aan 10:00.  
* Dan wordt \<Starttijd \- 15 min\> in de mail automatisch 09:45.  
* En wordt \<Starttijd \+ 60 min\> in de mail automatisch 11:00.

### 4.5.8	Logging en Foutafhandeling

Het mailmerge-systeem houdt gedetailleerde logs bij:

* **Per deelnemer:** Elke verzonden mail wordt gelogd met ontvanger, onderwerp, en toegevoegde bijlagen (zowel generiek als persoonlijk)
* **Foutmeldingen:** Als een mail niet kan worden verzonden (bijv. ongeldig e-mailadres), wordt de fout gelogd en gaat het proces door met de volgende deelnemer
* **Samenvatting:** Na afloop krijgt u een melding met het totaal aantal verzonden emails
* **Account verificatie:** Alle account- en aliaschecks worden gelogd voor controle

Deze logs zijn te vinden in het gedeelde logdocument (LOG_DOCUMENT_ID).

## 4.6	Importeren van Deelnemers via Excel

Voor besloten clinics kunnen deelnemers direct vanuit een Excel-bestand worden geïmporteerd.

### 4.6.1	Procedure

1. Plaats het Excel-bestand (.xlsx) in de daarvoor bestemde Google Drive map (EXCEL\_IMPORT\_FOLDER\_ID).  
2. Ga in het spreadsheet naar Thermoclinics Tools \-\> Lees Excel-bestand in voor besloten clinic.  
3. Selecteer het juiste bestand in het dialoogvenster en klik op "Importeer".  
4. Het systeem verwerkt het bestand. Na afloop verschijnt een samenvatting van het aantal toegevoegde, bijgewerkte en mislukte rijen.

### 4.6.2	Vereisten voor het Excel-bestand

Het systeem is flexibel in de exacte kolomnamen. Het zoekt naar de volgende koppen (hoofdletterongevoelig):

**Verplichte kolommen:**

* **Datum:** datum  
* **Tijd:** tijd  
* **Locatie:** locatie  
* **Voornaam:** voornaam  
* **Achternaam:** achternaam  
* **Email:** email of communications email address

**Optionele kolommen:**

* **Telefoonnummer:** telefoonnummer  
* **Geboortedatum:** geboortedatum  
* **Woonplaats:** woonplaats

**Belangrijk:** De combinatie van datum, tijd en locatie in het Excel-bestand moet exact overeenkomen met een clinic die al is gedefinieerd in de Data clinics sheet. Anders kan de deelnemer niet worden geplaatst.

### 4.6.3	Verwerking van Bestaande Deelnemers

Het systeem voorkomt dubbele inschrijvingen door intelligente detectie:

* **Identificatie:** Deelnemers worden geïdentificeerd via de combinatie van e-mailadres EN eventnaam
* **Bij bestaande deelnemer:** Als een deelnemer uit het Excel-bestand (op basis van e-mailadres) al voor **exact dezelfde clinic** staat geregistreerd:
  - De bestaande deelnemersgegevens worden **bijgewerkt** met de informatie uit het Excel-bestand
  - Voornaam, achternaam, telefoon, geboortedatum, en woonplaats worden geüpdatet indien aanwezig in het Excel-bestand
  - Het deelnemernummer en de Drive-map blijven **ongewijzigd**
  - Het aantal boekingen wordt **niet** verhoogd
  - De participant folder wordt automatisch hernoemd als de naam is gewijzigd
* **Bij nieuwe deelnemer:** 
  - Een nieuw deelnemernummer wordt toegewezen
  - Het aantal boekingen wordt verhoogd
  - Een nieuwe Drive-map wordt aangemaakt

**Workflow Excel + Formulier:**
1. Deelnemers worden eerst geïmporteerd via Excel (alleen email + basisgegevens)
2. Deelnemers vullen later zelf het formulier in met aanvullende informatie
3. Het systeem herkent de duplicate inschrijving en voegt de formuliergegevens toe aan de bestaande Excel-entry
4. De dubbele formulierinzending wordt automatisch verwijderd
5. Resultaat: Eén complete deelnemersrecord met zowel Excel- als formuliergegevens


