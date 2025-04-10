import os
import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox

# funksjon for å kombinere alle de 24 data filene inn til en DataFrame
def kombiner_excel_filer(input_dir='.'):
    
    """ definere excel_filer som valgt mappe i os.listdir om de er filtype .xlsx 
        if not: melding om feil, og returnerer Dataframe"""
    excel_filer = [fil for fil in os.listdir(input_dir) if fil.endswith('.xlsx')]
    if not excel_filer:
        messagebox.showinfo(
            title="Feil oppsto",
            message="Ingen Excel-filer funnet i valgt mappe."
        )
        return pd.DataFrame()  # Returner tom DataFrame

# lager tomt DataFrame forså å kombinerer dataen fra de ulike excel filene som er valgt
    kombinert_df = pd.DataFrame()
    for fil in excel_filer:
        filbane = os.path.join(input_dir, fil)
        # prøv å les filene some er valgt i filbane
        try:
            df = pd.read_excel(filbane, engine='openpyxl')
            kombinert_df = pd.concat([kombinert_df, df], ignore_index=True)
        except Exception as e:
            messagebox.showinfo(
                title="Feil oppdaget",
                message=f"Feil ved lesing av {fil}:\n{e}"
            )
            return pd.DataFrame()
    return kombinert_df

# funksjon for å erstate Not-a-number (NaN) verdier til 0 slik at det ikke oppstår noen feil ved plottingen. 
def rens_data_for_plotting(df):
  
    df.replace(r'^\s*$', np.nan, regex=True, inplace=True)
    renset_df = df.fillna(0)
    return renset_df

""" i sammedrag filen vår er kun ute etter å se på tider hvor det er aktivtet, så for å få færre 
    linjer i tabellen så fjerner vi rader som 0 i verdi i "Power [kW]" kolonnen"""
def filter_summary_data(df):
    
     #om man velger en mappe med excel filer som ikke har de radene vi ser etter.
    if 'Power [kW]' not in df.columns:
        messagebox.showinfo(
            title="Feil",
            message="Kolonnen 'Power [kW]' finnes ikke i dataene. " 
            "\nSjekk at de valgte filene har riktig struktur."
        )
        sys.exit(0)
    filtrert_df = df[df['Power [kW]'] != 0]
    return filtrert_df

# funksjon for plottingen
def plot_power_data(data):
    
    # definere varibaler fra datasettet  
    dato_tids_kolonne = 'Timestamp'
    y_data_kolonne = 'Power [kW]'

    # om excel filen ikke inneholder dato_tids_kolonnen
    if dato_tids_kolonne not in data.columns or y_data_kolonne not in data.columns:
        messagebox.showinfo(
            title="Feil",
            message=f"Nødvendige kolonner ('{dato_tids_kolonne}' og '{y_data_kolonne}') finnes ikke i dataen."
        )
        sys.exit(0)

    try:
        data[dato_tids_kolonne] = pd.to_datetime(data[dato_tids_kolonne], format='%d.%m.%Y %H:%M')
    except Exception as e:
        messagebox.showinfo(
            title="Feil",
            message=f"Feil ved konvertering av '{dato_tids_kolonne}':\n{e}"
        )
        sys.exit(0)

    # Ekstraher tid og dato da orginal data har tiddstempel dd.mm.yyyy HH:MM
    data['Time'] = data[dato_tids_kolonne].dt.strftime('%H:%M')
    data['Date'] = data[dato_tids_kolonne].dt.date

    # Fjerner duplikater slik at pivot tabellen ikke feiler
    data = data.drop_duplicates(subset=['Time', 'Date'])
    #Siden daten har en frekvent på hver 5. minut så gjør vi om til å kun beholde hele timer (frekvens til hver 10. minutt)
    data = data[data['Time'].str.match(r'^\d{2}:[0-5]0$')]
    
    # formatere dataseette til en "pivot tabell" 2D array  for å enklere beregne gjennomsnittet og for å visualiserer plottingen
    pivot_table = data.pivot(index='Time', columns='Date', values=y_data_kolonne)
    
    # nomenell effektfaktor for å beregne tilsynelatende effekt
    effektfaktor = 0.8648
    pivot_table_med_tap = pivot_table / effektfaktor

    # vektorisk beregning for verdiene som plottes ut. 
    gjennomsnitt_med_tap = pivot_table_med_tap.mean(axis=1)
    gjennomsnitt_med_tap = np.maximum(gjennomsnitt_med_tap, 0)
    max_med_tap = pivot_table_med_tap.max(axis=1)

    # oppretter figuren som plottes på
    fig, ax = plt.subplots(figsize=(10, 6))
    # prikkene som plottes på figuren, merketype 'o', spørrelse 2 og gjennomsiktighet på 50%
    for column in pivot_table.columns:
        ax.plot(pivot_table.index, pivot_table[column], 'o', markersize=2, alpha=0.5)

    # linje for gjennomsnitt, og peak i hver sin farge og linjestil
    ax.plot(gjennomsnitt_med_tap.index, gjennomsnitt_med_tap.values, label='Average power',
            color='red', linewidth=2)
    ax.plot(max_med_tap.index, max_med_tap.values, label='Max Total Power',
            color='blue', linewidth=2, linestyle='--')
    
    # navngir aksene, samt tittelen på grafen
    ax.set_xlabel('Time of day (hh:mm)')
    ax.set_ylabel('Power [kW]')
    ax.set_title('Power with loss over 24h')
    # legende for å forklare hva linjene visualiserer
    ax.legend()
    #synlig rutenett
    ax.grid(True)
    # setter hvor x aksen skal starte og slutte
    ax.set_xlim('00:00', '23:59')
    # setter hvor y aksen skal starte og slutte
    ax.set_ylim(bottom=0)
    # Velger å kun vise hver sjette interval, da dataen har en frekvens på hvert tiende minutt, 
    # slik at X aksen kun har synlige tall hver time
    ax.set_xticks(pivot_table.index[::6])
    # Rotere teksen langs x-aksen med 45 grader. 
    ax.set_xticklabels(pivot_table.index[::6], rotation=45)
    
    
    plt.show()

 # funksjon for hovedlogikken 
def main():

    while True:
        root = tk.Tk()
        root.withdraw()

        # bruker askokcancel slik at sluttbruker kan lukke programmet ved hjelp av X - knappen i vinduet.
        if not messagebox.askokcancel("Velkommen", 
            "Vennligst velg mappen som inneholder Excel-filer med relevant data."):
            sys.exit(0)
        
        #valgt å gå til sys.exit(0) fordi tidlig så havnet man i et vindu-loop som var uheldig.
        input_utforsker = filedialog.askdirectory(title="Velg katalog med Excel-filer")
        if not input_utforsker:
            if not messagebox.askokcancel("Ingen mappe valgt", "Ingen mappe valgt. Vil du prøve igjen?"):
                sys.exit(0)
            else:
                continue
        #printer til termimalen er kun for å hjelpe utvikler for å se at ting beveger seg fremover
        print(f"Du valgte mappen: {input_utforsker}")

        excel_filer = [fil for fil in os.listdir(input_utforsker) if fil.endswith('.xlsx')]
        if not excel_filer:
            if not messagebox.askokcancel("Ingen Excel-filer funnet", "Ingen Excel-filer funnet i valgt mappe. Vil du prøve igjen?"):
                sys.exit(0)
            else:
                continue
        else:
            if not messagebox.askokcancel("Filer funnet", "Filer funnet! Starter nå med komprimering."):
                sys.exit(0)

        kombinert_df = kombiner_excel_filer(input_utforsker)
        if kombinert_df.empty:
            if not messagebox.askokcancel("Feil oppsto", "Data fra Excel-filene er tomt. Vil du prøve igjen?"):
                sys.exit(0)
            else:
                continue
        break

    # Datastrøm for plotting: behold alle rader, erstatt NaN med 0
    plot_data_df = rens_data_for_plotting(kombinert_df.copy())
    # Datastrøm for Summary: filtrer ut rader med 0 i 'Power [kW]'
    summary_df = filter_summary_data(kombinert_df.copy())

    # Opprett "Summary"-mappen i den valgte mappen (input_utforsker)
    summary_mappe = os.path.join(input_utforsker, "Summary")
    os.makedirs(summary_mappe, exist_ok=True)
    summary_fil = os.path.join(summary_mappe, 'Summary.xlsx')

    try:
        summary_df.to_excel(summary_fil, index=False)
        #print er egentlig overflødig, men det er en indikator for at koden jobber. 
        print(f"Sammendragsfilen er lagret i mappen: {summary_mappe}")
    except Exception as e:
        if not messagebox.askokcancel("Feil ved lagring", f"Feil ved lagring av {summary_fil}:\n{e}\nVil du prøve igjen?"):
            sys.exit(0)
        else:
            sys.exit(0)

    # Plotting basert på datasettet med alle verdier (inkludert null)
    plot_power_data(plot_data_df)

#om __navn__ er 'main' så kjøres main funkjson, dette sørger for at programmet kjøres når skriptet kjøres. 
if __name__ == '__main__':
    main()
