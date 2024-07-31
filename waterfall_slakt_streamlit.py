# -*- coding: utf-8 -*-
"""
Created on Wed Jul 31 08:53:23 2024

@author: HåkonTveiten
"""
import pandas as pd
from datetime import datetime
import numpy as np
import matplotlib.pyplot as plt
import streamlit as st

def les_data(uploaded_file):
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, header=2)
            return df
        except Exception as e:
            st.error(f"Feil ved lesing av Excel-filen: {e}")
            return None
    return None

def beregn_stopptid(row, sheet_type):
    try:
        if sheet_type == "slakt":
            stopptid = (
                row.iloc[27:31].fillna(0).sum() +
                row.iloc[34:40].fillna(0).sum() / 6 +
                row.iloc[40:51].fillna(0).sum()
            )
        elif sheet_type == "filet":
            stopptid = (
                row.iloc[32:40].fillna(0).sum() +
                row.iloc[40:52].fillna(0).sum()
            )
        return stopptid
    except Exception as e:
        st.error(f"Feil ved beregning av stopptid: {e}")
        return None

def beregn_faktiskproduksjon(row, sheet_type):
    try:
        if sheet_type == "slakt":
            arbeidstimer = (datetime.strptime(str(row.iloc[3]), "%H:%M:%S") - datetime.strptime(str(row.iloc[2]), "%H:%M:%S")).seconds / 3600
            arbeidstimer = arbeidstimer * 60
            antall_fisk = row.iloc[4]
        elif sheet_type == "filet":
            arbeidstimer = (datetime.strptime(str(row.iloc[7]), "%H:%M:%S") - datetime.strptime(str(row.iloc[6]), "%H:%M:%S")).seconds / 3600
            arbeidstimer = arbeidstimer * 60
            antall_fisk = row.iloc[12]
        return arbeidstimer, antall_fisk
    except Exception as e:
        st.error(f"Feil ved beregning av faktisk produksjon: {e}")
        return None, None

def velg_dato():
    år = st.number_input("Tast inn året du ønsker å sjekke:", min_value=2024, max_value=datetime.now().year)
    maneder = ["Januar", "Februar", "Mars", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Desember"]
    maaned = st.selectbox("Velg måned:", list(range(1, 13)), format_func=lambda x: maneder[x-1])
    dag = st.number_input("Tast inn dagen i måneden (1-31):", min_value=1, max_value=31)
    valgt_dato = datetime(år, maaned, dag)
    return valgt_dato

def main():
    st.title("Produksjonsanalyse")
    
    # Velg type ark
    sheet_type = st.selectbox("Velg type ark:", ["slakt", "filet"])
    oee_100 = 150 if sheet_type == "slakt" else 25
    stiplet_hoeyde = 120 if sheet_type == "slakt" else 20
    
    # Filopplastingsseksjon
    uploaded_file = st.file_uploader(f"Velg en Excel-fil (må være et 'input{sheet_type}'-ark).", type=["xlsx"])
    
    if uploaded_file is None:
        st.warning("Vennligst last opp en Excel-fil for å fortsette.")
        return

    # Last inn data fra opplastet fil
    df = les_data(uploaded_file)
    
    if df is None:
        st.warning("Ingen data tilgjengelig i den opplastede filen. Vennligst last opp en gyldig Excel-fil.")
        return

    valgt_dato = velg_dato()

    # Konverter første kolonne til datetime-format
    try:
        df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], format="%Y-%m-%d %H:%M:%S")
    except ValueError as e:
        st.error(f"Feil ved konvertering av dato: {e}")
        st.write("Det kan være et problem med datoformatet i filen, eller du har lastet opp feil type fil.")
        return
    
    # Sjekk om valgt dato finnes i første kolonne
    st.write(f"Valgt dato: {valgt_dato.date()}")
    st.write("----------------------------------")
    st.write("Statistikk")
    
    # Filtrer datoene for å sammenligne basert på år, måned og dag
    try:
        df['Dato'] = df.iloc[:, 0].dt.date
    except Exception as e:
        st.error(f"Feil ved behandling av datoer: {e}")
        st.write("Det kan være et problem med strukturen på den opplastede filen.")
        return

    valgt_dato_enkel = valgt_dato.date()
    
    if valgt_dato_enkel in df['Dato'].values:
        row = df[df['Dato'] == valgt_dato_enkel].iloc[0]

        stopptid = beregn_stopptid(row, sheet_type)
        if stopptid is None:
            st.error("Kan ikke beregne stopptid. Sjekk om du har valgt riktig ark type (file eller slakt) og lastet opp riktig fil.")
            return
        
        stopptid_impact = stopptid * oee_100
        stopptid_takt = round(stopptid_impact / 60 / 8, 2)
        
        arbeidstimer, antall_fisk = beregn_faktiskproduksjon(row, sheet_type)
        if arbeidstimer is None or antall_fisk is None:
            st.error("Kan ikke beregne faktisk produksjon. Sjekk om du har valgt riktig filtype og lastet opp riktig fil.")
            return
        
        # Define the width for alignment
        label_width = 70  # Adjust this value based on your label lengths
        value_width = 10  # Adjust this value based on your value lengths

        # Format the output with padding
        st.write(f'OEE 100%:'.ljust(label_width) + f'{oee_100:<{value_width}} fisk/minutt')
        st.write(f'Total stopptid:'.ljust(label_width) + f'{round(stopptid, 2):<{value_width}} minutter')
        st.write(f'Tapt takt per minutt på grunn av stopp:'.ljust(label_width) + f'{stopptid_takt:<{value_width}} fisk/minutt')
        st.write("")
        st.write(f'Arbeidstimer:'.ljust(label_width) + f'{round(arbeidstimer/60, 2):<{value_width}} timer')
        st.write(f'Antall fisk produsert:'.ljust(label_width) + f'{antall_fisk:<{value_width}} fisk')
        faktisk_takt = round(antall_fisk / arbeidstimer, 2)

        kjente_faktorer = round(stopptid_takt, 2)
        annet = oee_100 - kjente_faktorer - faktisk_takt
        annet = round(annet, 2)

        st.write(f'Antall fisk tapt pga stopptid:'.ljust(label_width) + f'{round(stopptid_impact, 2):<{value_width}} fisk')
        st.write(f'Annet tap (unoterte feil, operatørhastighet etc):'.ljust(label_width) + f'{annet:<{value_width}} minutter')
        st.write("")
        st.write(f'Faktisk takt:'.ljust(label_width) + f'{faktisk_takt:<{value_width}} fisk/minutt')
        st.write(f'Avstand fra 80% OEE takttid:'.ljust(label_width) + f'{round(stiplet_hoeyde - faktisk_takt, 2):<{value_width}} fisk/minutt')


        # Data for waterfall grafen
        stages = ['100% OEE', 'Stopptid', 'Annet']
        values = [oee_100, -stopptid_takt, -annet]

        cum_values = np.cumsum([0] + values).tolist()
        value_starts = cum_values[:-1]

        # Plot waterfall grafen
        fig, ax = plt.subplots()
        colors = ['blue', 'red', 'orange']

        for i in range(len(stages)):
            ax.bar(stages[i], values[i], bottom=value_starts[i], color=colors[i], edgecolor='black')

        ax.bar('Takttid', faktisk_takt, bottom=0, color='green', edgecolor='black')
        ax.bar('Takttid', stiplet_hoeyde - faktisk_takt, bottom=faktisk_takt, color='none', edgecolor='green', hatch='//')

        for i in range(len(stages)):
            y_pos = value_starts[i] + values[i] / 2
            ax.text(stages[i], y_pos, f'{values[i]}', ha='center', va='center', color='white', fontweight='bold')

        ax.text('Takttid', faktisk_takt / 2, f'{faktisk_takt}', ha='center', va='center', color='white', fontweight='bold')
        ax.text('Takttid',faktisk_takt + (stiplet_hoeyde - faktisk_takt) / 2, f'{round(stiplet_hoeyde - faktisk_takt, 2)}', ha='center', va='center', color='green', fontweight='bold')

        ax.set_ylabel('Produksjonsverdi')
        ax.set_title(f'Produksjon for {valgt_dato_enkel} ({sheet_type})')
        st.pyplot(fig)

    else:
        st.warning("Datoen du valgte finnes ikke i input-arket. Dette er enten fordi du tastet inn en ugyldig dato eller fordi datoen ikke hadde noen produksjon (eks helg).")

if __name__ == "__main__":
    main()
