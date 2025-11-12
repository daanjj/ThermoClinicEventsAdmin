# Cancelen van deelnemers of verschuiven naar andere datum

Het komt met enige regelmaat voor dat een deelnemer zich terugtrekt of bij nader inzien een clinic op een andere datum wil volgen. Hoe hiermee om te gaan?

### Wat gebeurt er als iemand zich aanmeldt of als een Excel-file wordt ingelezen?

Als een deelnemer via het systeem voor een event wordt geregistreerd, dan gebeurt het volgende automatisch:

1. De deelnemer krijgt een \<deelnemernummer\> toegewezen; als er al 2 boekingen zijn voor het event wordt zijn/haar deelnemernummer 3 (het deelnemernummer is vooral handig bij de toewijzing van sensoren en watches)  
2. Er wordt een submap voor deze deelnemer aangemaakt in de map van het event. Deze map heeft als naam “\<Deelnemernummer\> \<Voornaam\> \<Achternaam\>”  
3. Het ‘ID’ van deze submap (een code die kan worden gebruikt om de map zonder zoeken snel terug te vinden) wordt toegevoegd aan de sheet ‘Open Form Responses’ of ‘Besloten Form Responses’. Dit ID wordt gebruikt bij het mailen van deelnemerspecifieke bestanden (FIT-files, certificaten, …)  
4. Het aantal boekingen voor dit event wordt met 1 verhoogd  
5. De pop-ups in de aanmeldformulieren worden aangepast (omdat in de pop-ups het aantal deelnemers staan én omdat een event uit de pop-up valt zodra het maximale aantal deelnemers is bereikt)   
6. De agenda-afspraak voor het event wordt aangepast (in de titel staat het aantal deelnemers, in het voorbeeld nu dus 3\)  
   

### Hoe een cancelation verwerken?

Als de deelnemer simpelweg zou worden verwijderd uit de deelnemerregistratie dan …:

- Blijft de submap voor deze deelnemer bestaan (heeft verder geen consequenties)  
- Klopt het aantal boekingen niet meer \-\> dat zou je handmatig met 1 kunnen verlagen maar… dan gaat het nummeren van deelnemers mis. Immers, als kandidaat 1 cancelt en er zijn 3 boekingen, dan zou het aantal boekingen op 2 gezet moeten worden. De eerstvolgende die zich daarna aanmeldt zou dan deelnemernummer 3 krijgen, een nummer dat al toegekend is aan iemand anders.  
- De pop-ups worden correct bijgewerkt als het deelnemernummer verlaagd wordt  
- Datzelfde geldt voor de agenda-afspraken 

Wat is beter:

1. **Verwijder de deelnemer** (of: maak in ieder geval het geselecteerde event leeg zodat de deelnemer niet meer aan een event is gekoppeld)  
2. **Verhoog het maximum aantal boekingen** voor dat event met 1, want dan:  
- Ontstaan er geen dubbele deelnemernummers wat handig is bij koppelen aan sensors en watches \-\> als de inschrijving is gesloten is het handig om het deelnemernummer van iemand met een ‘te hoog’ deelnemernummer (hoger dan het maximum) in de administratie aan te passen door ze handmatig een vrijgekomen deelnemernummer toe te kennen  
- De pop-ups die het aantal beschikbare plaatsen aangeven blijven kloppen want het maximum is met 1 verhoogd maar er is dus ook 1 ‘ongeldige’ deelnemer  
3. **Verander het volgnummer van de deelnemermap naar XX** (dus bv van ‘02 Joost Fonville’ naar ‘XX Joost Fonville’)   
   * Hierdoor wordt deze deelnemer niet meer meegenomen bij het opstellen van een deelnemerslijst  
   * Als je in de map van een event alle deelnemermaps op alfabet sorteert komen de mapjes van verwijderde deelnemers onderaan te staan. Je kan dan ook snel zien welke volgnummers nog ontbreken

Toewijzen lager volgnummer indien deelnemer een volgnummer heeft dat hoger is dan aantal beschikbare CORE-sensoren voor een clinic  
Als een deelnemer een volgnummer heeft gekregen dat hoger is dan het aantal beschikbare CORE-sensoren, bv deelnemer 16 terwijl er maar 15 sensoren zijn:

- Wijs een vrijgevallen nummer (‘02’ in het voorbeeld van de vorige paragraaf) toe aan een deelnemer met zo’n ‘te hoog’ volgnummer.   
- Dit doe je door de betreffende deelnemermap te hernoemen zodat die met het juiste volgnummer begint (dus in het voorbeeld hierboven: verander de eerste twee karakters van de deelnemermap van ‘16’ naar ‘02’). 

**Opmerking: agenda-afspraken worden bij een cancelation NIET aangepast**  
De agenda-afspraak zal bij een cancelation niet meer kloppen omdat in de agenda-afspraak het aantal aangemelde deelnemers wordt vermeld. Dat nummer wordt niet gecorrigeerd als een deelnemer zich heeft teruggetrokken en op de hiervoor beschreven wijze administratief wordt verwijderd.

### Hoe een deelnemer naar een andere datum doorschuiven

Simpelweg het event aanpassen voor de deelnemer gaat mis omdat:

- De deelnemer geen nieuw deelnemernummer toegekend krijgt  
- Er geen nieuwe deelnemermap wordt aangemaakt   
- De reeds aangemaakte submap nu bij het verkeerde event staat (wat het lastig maakt om zijn/haar FIT-files te mailen want waar staat die map dan??)  
- Het aantal boekingen voor het nieuwe event niet met 1 wordt verhoogd en voor het oude event niet met 1 wordt verlaagd (ze zijn nu dus allebei incorrect)  
- Ook de pop-ups en agenda-afspraken nu incorrect zijn

**Hoe dan wel een deelnemer doorschuiven?**  
‘Cancel’ de deelnemer voor het oude event (zie boven) én meldt hem/haar opnieuw ‘via de voordeur’ aan zodat de 6 hierboven genoemde zaken allemaal worden uitgevoerd. Dus:

- Laat de deelnemer zichzelf opnieuw aanmelden via het Open of het Besloten formulier, óf  
- Neem de deelnemer in een Excel op en lees die Excel in voor de betreffende clinic..