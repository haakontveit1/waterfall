# -*- coding: utf-8 -*-
"""
Created on Wed Jul 31 08:53:23 2024

@author: HåkonTveiten
"""

import pandas as pd
from datetime import datetime
import numpy as np
import matplotlib.pyplot as plt
import os
import streamlit as st

def les_data():
    # Få stien til den katalogen som inneholder dette skriptet
    script_dir = os.path.dirname(os.path.realpath(__file__))
    # Bygg stien til Excel-filen
    file_path = os.path.join(script_dir, "excelark", "inputslakt2907.xlsx")
    # Les data fra Excel-filen
    df = pd.read_excel(file_path, header=2)
    return df

def beregn_stopptid(row):
    # Kalkuler stopptid ved bruk av kolonneindekser (0-indeksert)
    stopptid = (
        row.iloc[27] + row.iloc[28] + row.iloc[29] + row.iloc[30] +  # AB til AE (kolonneindekser 27-30)
        (row.iloc[34] + row.iloc[35] + row.iloc[36] + row.iloc[37] + row.iloc[38] + row.iloc[39]) / 6 +  # AI til AN (kolonneindekser 34-39)
        row.iloc[40] + row.iloc[41] + row.iloc[42] + row.iloc[43] + row.iloc[44] + row.iloc[45] + row.iloc[46] + row.iloc[47] + row.iloc[48] + row.iloc[49] + row.iloc[50]  # AO til AY (kolonneindekser 40-50)
    )
    return stopptid

def beregn_faktiskproduksjon(row): 
    # Kalkuler faktisk produksjon ved bruk av kolonneindekser (0-indeksert)
    arbeidstimer = (datetime.strptime(str(row.iloc[3]), "%H:%M:%S") - datetime.strptime(str(row.iloc[2]), "%H:%M:%S")).seconds / 3600  # D (kolonne 3) - C (kolonne 2)
    arbeidstimer = arbeidstimer * 60
    antall_fisk = row.iloc[4]  # E (kolonne 4)
    return arbeidstimer, antall_fisk

def velg_dato():
    # Spør bruker om år, måned og dag
    år = st.number_input("Tast inn året du ønsker å sjekke:", min_value=2000, max_value=datetime.now().year)
    maneder = ["Januar", "Februar", "Mars", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Desember"]
    maaned = st.selectbox("Velg måned:", list(range(1, 13)), format_func=lambda x: maneder[x-1])
    dag = st.number_input("Tast inn dagen i måneden (1-31):", min_value=1, max_value=31)
    valgt_dato = datetime(år, maaned, dag)
    return valgt_dato

def main():
    st.title("Produksjonsanalyse")
    oee_100 = 150
    stiplet_hoeyde = 120
    
    df = les_data()
    valgt_dato = velg_dato()

    # Konverter første kolonne til datetime-format
    try:
        df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], format="%Y-%m-%d %H:%M:%S")
    except ValueError as e:
        st.error(f"Feil ved konvertering av dato: {e}")
        st.write("Første kolonne verdier:", df.iloc[:, 0].head())
        return
    
    # Sjekk om valgt dato finnes i første kolonne
    st.write(f"Valgt dato: {valgt_dato.date()}")
    st.write("----------------------------------")
    st.write("Statistikk")
    
    # Filtrer datoene for å sammenligne basert på år, måned og dag
    df['Dato'] = df.iloc[:, 0].dt.date
    valgt_dato_enkel = valgt_dato.date()
    
    if valgt_dato_enkel in df['Dato'].values:
        row = df[df['Dato'] == valgt_dato_enkel].iloc[0]

        stopptid = beregn_stopptid(row)
        stopptid_impact = stopptid * oee_100
        stopptid_takt = round(stopptid_impact / 60 / 8, 2)
        
        arbeidstimer, antall_fisk = beregn_faktiskproduksjon(row)
        
        st.write(f'OEE 100%: {oee_100}')
        st.write(f'Total stopptid: {round(stopptid, 2)}')
        st.write(f'Tapt takt per minutt på grunn av stopp: {stopptid_takt}')
        st.write("")
        st.write(f'Arbeidstimer: {round(arbeidstimer/60,2)}')
        st.write(f'Antall fisk produsert: {antall_fisk}')
        faktisk_takt = round(antall_fisk / arbeidstimer, 2)
        
        # Beregn ukjent faktor "Annet"
        kjente_faktorer = round(stopptid_takt, 2)
        annet = oee_100 - kjente_faktorer - faktisk_takt
        annet = round(annet, 2)

        # Print verdiene for de forskjellige kolonnene
        st.write(f'Antall fisk tapt pga stopptid: {round(stopptid_impact,2)}')
        st.write(f'Annet tap (unoterte feil, operatørhastighet etc): {annet}')
        st.write("")
        st.write(f'Faktisk takt: {faktisk_takt}')
        st.write(f'Avstand fra 80% OEE takttid: {round(stiplet_hoeyde - faktisk_takt, 2)}')

        # Data for waterfall grafen
        stages = ['100% OEE', 'Stopptid', 'Annet']
        values = [oee_100, -stopptid_takt, -annet]

        # Beregn mellomverdier
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
        ax.set_title(f'Produksjon for {valgt_dato_enkel}')
        st.pyplot(fig)

    else:
        st.warning("Datoen du valgte finnes ikke i input-arket. Dette er enten fordi du tastet inn en ugyldig dato eller fordi datoen ikke hadde noen produksjon (eks helg).")

if __name__ == "__main__":
    main()
