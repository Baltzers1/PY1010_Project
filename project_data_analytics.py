import os
import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox

def combine_excel_files(input_dir='.'):
    """
    Kombinerer alle Excel-filer (.xlsx) i angitt mappe til en enkelt DataFrame.
    Forutsetter at alle filene har samme datastruktur.

    Parameter:
      input_dir (str): Mappen der Excel-filene befinner seg.
      
    Returnerer:
      Pandas DataFrame med alle innleste data.
    """
    excel_files = [file for file in os.listdir(input_dir) if file.endswith('.xlsx')]
    if not excel_files:
        messagebox.showinfo(
            title="Feil oppsto",
            message="Ingen Excel-filer funnet i valgt mappe."
        )
        return pd.DataFrame()  # Returner tom DataFrame

    combined_df = pd.DataFrame()
    for file in excel_files:
        file_path = os.path.join(input_dir, file)
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            combined_df = pd.concat([combined_df, df], ignore_index=True)
        except Exception as e:
            messagebox.showinfo(
                title="Feil oppdaget",
                message=f"Feil ved lesing av {file}:\n{e}"
            )
            return pd.DataFrame()
    return combined_df

def clean_data_for_plotting(df):
    """
    Renser data ved å erstatte NaN-verdier med 0, slik at alle radene beholdes.
    Datasettet brukes til utregning av gjennomsnitt og plotting, slik at perioder uten aktivitet (0) også tas med.
    
    Parameter:
      df (DataFrame): DataFrame med de kombinerte dataene.
      
    Returnerer:
      DataFrame der NaN-verdier er erstattet med 0.
    """
    df.replace(r'^\s*$', np.nan, regex=True, inplace=True)
    clean_df = df.fillna(0)
    return clean_df

def filter_summary_data(df):
    """
    Filtrerer bort rader der 'Power [kW]' er 0, slik at sammendragsfilen kun inneholder
    rader med reell aktivitet.
    
    Parameter:
      df (DataFrame): DataFrame med de kombinerte dataene.
      
    Returnerer:
      DataFrame der rader med 0 i 'Power [kW]' er filtrert bort.
    """
    if 'Power [kW]' not in df.columns:
        messagebox.showinfo(
            title="Feil",
            message="Kolonnen 'Power [kW]' finnes ikke i dataene. "
                    "Sjekk at de valgte filene har riktig struktur."
        )
        sys.exit(0)
    filtered_df = df[df['Power [kW]'] != 0]
    return filtered_df

def plot_power_data(data):
    """
    Plotter data for et 24-timers vindu basert på et komplett datasett (inkludert nullverdier).
    Dataen benyttes til å beregne gjennomsnitt og peak, slik at perioder med nullaktivitet er med i utregningen.
    
    Parameter:
      data (DataFrame): Datasettet som skal plottes.
    """
    date_time_column = 'Timestamp'
    y_data_column = 'Power [kW]'

    if date_time_column not in data.columns or y_data_column not in data.columns:
        messagebox.showinfo(
            title="Feil",
            message=f"Nødvendige kolonner ('{date_time_column}' og '{y_data_column}') finnes ikke i dataen."
        )
        sys.exit(0)

    try:
        data[date_time_column] = pd.to_datetime(data[date_time_column], format='%d.%m.%Y %H:%M')
    except Exception as e:
        messagebox.showinfo(
            title="Feil",
            message=f"Feil ved konvertering av '{date_time_column}':\n{e}"
        )
        sys.exit(0)

    # Ekstraher tid og dato
    data['Time'] = data[date_time_column].dt.strftime('%H:%M')
    data['Date'] = data[date_time_column].dt.date

    # Fjern duplikater og filtrer tidspunkter med riktig format
    data = data.drop_duplicates(subset=['Time', 'Date'])
    data = data[data['Time'].str.match(r'^\d{2}:[0-5]0$')]

    pivot_table = data.pivot(index='Time', columns='Date', values=y_data_column)
    
    efficiency_factor = 0.8648
    pivot_table_with_loss = pivot_table / efficiency_factor

    average_with_loss = pivot_table_with_loss.mean(axis=1)
    average_with_loss = np.maximum(average_with_loss, 0)
    max_with_loss = pivot_table_with_loss.max(axis=1)

    fig, ax = plt.subplots(figsize=(10, 6))
    for column in pivot_table.columns:
        ax.plot(pivot_table.index, pivot_table[column], 'o', markersize=2, alpha=0.5)

    ax.plot(average_with_loss.index, average_with_loss.values, label='Average power',
            color='red', linewidth=2)
    ax.plot(max_with_loss.index, max_with_loss.values, label='Max Total Power',
            color='blue', linewidth=2, linestyle='--')

    ax.set_xlabel('Time of day (hh:mm)')
    ax.set_ylabel('Power [kW]')
    ax.set_title('Power with loss over 24h')
    ax.legend()
    ax.grid(True)
    ax.set_xticks(pivot_table.index[::6])
    ax.set_xticklabels(pivot_table.index[::6], rotation=45)
    
    plt.tight_layout()
    plt.show()

def main():
    """
    Hovedfunksjonen:
      1. Viser en "Velkommen"-melding.
      2. Lar brukeren velge mappen via filutforsker.
      3. Sjekker om mappen inneholder Excel-filer.
      4. Kombinerer Excel-filene til én DataFrame.
      5. Lager to datastrømmer:
         - Ett datasett for plotting (der NaN erstattes med 0, slik at alle perioder med aktivitet og inaktivitet er med).
         - Ett filtrert datasett (Summary) der rader med 0 i 'Power [kW]' fjernes.
      6. Skriver Summary til fil og plotter data basert på datasettet for plotting.
      
      Dersom brukeren klikker X‑knappen i et dialogvindu (eller trykker Cancel ved askokcancel),
      avsluttes programmet.
    """
    while True:
        root = tk.Tk()
        root.withdraw()

        # Vis "Velkommen"-melding med askokcancel
        if not messagebox.askokcancel("Velkommen", 
            "Vennligst velg mappen som inneholder Excel-filer med relevant data."):
            sys.exit(0)

        input_directory = filedialog.askdirectory(title="Velg katalog med Excel-filer")
        if not input_directory:
            messagebox.askokcancel("Ingen mappe valgt", "Ingen mappe valgt. Programmet lukkes")
            sys.exit(0)
         

        excel_files = [file for file in os.listdir(input_directory) if file.endswith('.xlsx')]
        if not excel_files:
            if not messagebox.askokcancel("Ingen Excel-filer funnet", "Ingen Excel-filer funnet i valgt mappe. Vil du prøve igjen?"):
                sys.exit(0)
            else:
                continue
        else:
            if not messagebox.askokcancel("Filer funnet", "Filer funnet! Starter nå med komprimering."):
                sys.exit(0)

        combined_df = combine_excel_files(input_directory)
        if combined_df.empty:
            if not messagebox.askokcancel("Feil oppsto", "Data fra Excel-filene er tomt. Vil du prøve igjen?"):
                sys.exit(0)
            else:
                continue
        break

    # Datastrøm for plotting: behold alle rader, erstatt NaN med 0
    plot_data_df = clean_data_for_plotting(combined_df.copy())
    # Datastrøm for Summary: filtrer ut rader med 0 i 'Power [kW]'
    summary_df = filter_summary_data(combined_df.copy())
    summary_file = 'Summary.xlsx'
    try:
        summary_df.to_excel(summary_file, index=False)
        print(f"Sammendragsfilen er lagret som '{summary_file}'.")
    except Exception as e:
        messagebox.askokcancel("Feil ved lagring", f"Feil ved lagring av {summary_file}:\n{e}")
        sys.exit(0)

    # Plotting basert på datasettet med alle verdier (inkludert null)
    plot_power_data(plot_data_df)

if __name__ == '__main__':
    main()
