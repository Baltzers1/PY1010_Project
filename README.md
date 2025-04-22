# Project: Data analytics

Dette programmet kombinerer, filtrerer og visualiserer effektdata hentet fra Excel-filer. Brukeren velger selv en mappe med .xlsx-filer,
og programmet genererer både en graf og en sammendragsfil (Summary.xlsx) basert på strømforbruk (Power [kW]) gjennom døgnet.

Last ned og pakk ut data:
I repository-et finner du en .zip-fil.
Last ned og pakk ut denne lokalt på din maskin.
Dette vil gi deg en mappe med 24 Excel-filer (.xlsx) som inneholder timesbasert effektforbruk.

Kjør programmet: 
Sørg for at du har installert nødvendige pakker (se nedenfor), og kjør programmet fra terminal eller IDE:
python PY1010_project.py

OBS: Du vil bli bedt om å velge mappen som inneholder .xlsx-filene via et grafisk grensesnitt (Tkinter).

## Hva programmet gjør:
1. Laster inn og kombinerer alle .xlsx-filer fra valgt mappe.
1. Renser data ved å erstatte manglende eller tomme verdier med 0.
1. Filtrerer vekk rader hvor 'Power [kW]' er lik 0, for å lage en kortere og mer informativ oppsummering.
1. Lagrer et sammendrag (Summary.xlsx) i en egen Summary-mappe.
1. Visualiserer alle registrerte effektbruk over et døgn:
1. Plottet viser alle målinger som prikker.
1. En rød linje viser gjennomsnittlig strømbruk etter beregnet effektap.
1. En blå stiplet linje viser maksverdiene.

## Eksempel på output:
Plottet visualiserer variasjoner i effektforbruk time for time. Dette kan være nyttig for å identifisere:
Forbrukstopper
Gjennomsnittlig effektbruk i løpet av døgnet
Tidspunkter med lav aktivitet

![results](https://github.com/user-attachments/assets/564f7b76-8bb4-4689-a2f2-ccba9de5ec84)

## Programmet krever følgende Python-pakker:

pandas,
numpy,
matplotlib,
tkinter,
openpyxl

## Installer & kjør fra kommandolinje
```sh
git clone git@github.com:Baltzers1/PY1010_Project.git
cd PY1010_Project
pip install pandas numpy matplotlib openpyxl tkinter   # Anbefales å installere pip-pakker i virtualenv
unzip baltzers1_data.zip                               # Unzip .xlsx filer til DATA/-mappen
python project_data_analytics.py                       # Kjør og velg DATA/-mappen i filvelgeren som dukker opp
```

Tips og feilhåndtering:
Dersom du velger en mappe uten Excel-filer, eller filene mangler kolonnen "Power [kW]", vil du få en feilmelding.
Programmet stopper og lar deg velge mappe på nytt dersom noe går galt, uten å kræsje.
Dataene forventes å ha tidspunkter i formatet dd.mm.yyyy HH:MM.
