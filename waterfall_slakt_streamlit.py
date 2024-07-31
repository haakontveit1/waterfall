# -*- coding: utf-8 -*-
"""
Created on Wed Jul 31 08:53:23 2024

@author: HåkonTveiten
"""
import pandas as pd
from datetime import datetime, timedelta
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
    år = st.number_input("Velg år:", min_value=2024, max_value=datetime.now().year)
    maneder = ["Januar", "Februar", "Mars", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Desember"]
    maaned = st.selectbox("Velg måned:", list(range(1, 13)), format_func=lambda x: maneder[x-1])
    dag = st.number_input("Velg dag (1-31):", min_value=1, max_value=31)
    valgt_dato = datetime(år, maaned, dag)
    return valgt_dato

def hent_uke_dager(år, uke_nummer):
    start_dato = datetime.strptime(f'{år}-W{uke_nummer-1}-1', "%Y-W%W-%w")
    return [start_dato + timedelta(days=i) for i in range(5)]  # Returner mandag-fredag

def main():
    st.title("Produksjonsanalyse")
    
    fisk = "fisk"
    pa = " (på slakt)"

    # Velg type analyse
    sheet_type = st.selectbox("Velg type ark:", ["slakt", "filet"])
    if sheet_type == "filet":
        fisk = "fileter"
        pa = ""


    uploaded_file = st.file_uploader(f"Velg en Excel-fil (må være et 'input-{sheet_type}'-ark).", type=["xlsx"])
    analyse_type = st.selectbox("Velg analyse:", ["Spesifikk dato", "Gjennomsnitt for en uke"])
    oee_100 = 150 if sheet_type == "slakt" else 25
    stiplet_hoeyde = 120 if sheet_type == "slakt" else 20
    
    if analyse_type == "Spesifikk dato":
        valgt_dato = velg_dato()
    else:
        år = st.number_input("Velg år:", min_value=2024, max_value=datetime.now().year)
        uke_nummer = st.number_input("Velg uke nummer:", min_value=1, max_value=52)
        uker_dager = hent_uke_dager(år, uke_nummer)
    
    if uploaded_file is None:
        st.warning("Vennligst last opp en Excel-fil for å fortsette.")
        return

    df = les_data(uploaded_file)
    
    if df is None:
        st.warning("Ingen data tilgjengelig i den opplastede filen. Vennligst last opp en gyldig Excel-fil.")
        return

    # Konverter første kolonne til datetime-format
    try:
        df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], format="%Y-%m-%d %H:%M:%S")
    except ValueError as e:
        st.error(f"Feil ved konvertering av dato: {e}")
        st.write("Det kan være et problem med datoformatet i filen, eller du har lastet opp feil type fil.")
        return

    # Filtrer datoene for å sammenligne basert på år, måned og dag
    try:
        df['Dato'] = df.iloc[:, 0].dt.date
    except Exception as e:
        st.error(f"Feil ved behandling av datoer: {e}")
        st.write("Det kan være et problem med strukturen på den opplastede filen.")
        return

    if analyse_type == "Spesifikk dato":
        valgt_dato_enkel = valgt_dato.date()
        if valgt_dato_enkel in df['Dato'].values:
            row = df[df['Dato'] == valgt_dato_enkel].iloc[0]
            stopptid = beregn_stopptid(row, sheet_type)
            arbeidstimer, antall_fisk = beregn_faktiskproduksjon(row, sheet_type)
            if stopptid is None or arbeidstimer is None or antall_fisk is None:
                st.error("Kan ikke beregne verdier. Sjekk om du har valgt riktig filtype og lastet opp riktig fil.")
                return
            
            stopptid_impact = stopptid * oee_100
            stopptid_takt = round(stopptid_impact / 60 / 8, 2)
            faktisk_takt = round(antall_fisk / arbeidstimer, 2)
            kjente_faktorer = round(stopptid_takt, 2)
            annet = oee_100 - kjente_faktorer - faktisk_takt
            annet = round(annet, 2)
            
            # Plotting single day data
            fig, ax = plt.subplots()
            ax.bar(['100% OEE', 'Stopptid', 'Annet'], [oee_100, -stopptid_takt, -annet], color=['blue', 'red', 'orange'])
            ax.bar('Takttid', faktisk_takt, bottom=0, color='green', edgecolor='black')
            ax.bar('Takttid', stiplet_hoeyde - faktisk_takt, bottom=faktisk_takt, color='none', edgecolor='green', hatch='//')
            ax.set_ylabel(f'Antall {fisk} produsert per minutt')
            ax.set_title(f'Antall {fisk} produsert per minutt {valgt_dato.strftime("%d.%m.%Y")}{pa}')
            st.pyplot(fig)
        else:
            st.warning("Datoen du valgte finnes ikke i input-arket. Dette er enten fordi du tastet inn en ugyldig dato eller fordi datoen ikke hadde noen produksjon (eks helg).")
    else:
        daglig_data = []
        for dag in uker_dager:
            dag_enkel = dag.date()
            if dag_enkel in df['Dato'].values:
                row = df[df['Dato'] == dag_enkel].iloc[0]
                stopptid = beregn_stopptid(row, sheet_type)
                arbeidstimer, antall_fisk = beregn_faktiskproduksjon(row, sheet_type)
                if stopptid is None or arbeidstimer is None or antall_fisk is None:
                    st.error(f"Kan ikke beregne verdier for {dag.strftime('%d.%m.%Y')}. Sjekk om du har valgt riktig filtype og lastet opp riktig fil.")
                    return
                daglig_data.append((dag, stopptid, arbeidstimer, antall_fisk))

        if not daglig_data:
            st.warning("Ingen gyldige data funnet for den valgte uken.")
            return

        # Plotting for each day
        fig, axes = plt.subplots(2, 3, figsize=(15, 10))
        for i, (dag, stopptid, arbeidstimer, antall_fisk) in enumerate(daglig_data):
            ax = axes.flatten()[i]
            stopptid_impact = stopptid * oee_100
            stopptid_takt = round(stopptid_impact / 60 / 8, 2)
            faktisk_takt = round(antall_fisk / arbeidstimer, 2)
            kjente_faktorer = round(stopptid_takt, 2)
            annet = oee_100 - kjente_faktorer - faktisk_takt
            annet = round(annet, 2)
            
            ax.bar(['100% OEE', 'Stopptid', 'Annet'], [oee_100, -stopptid_takt, -annet], color=['blue', 'red', 'orange'])
            ax.bar('Takttid', faktisk_takt, bottom=0, color='green', edgecolor='black')
            ax.bar('Takttid', stiplet_hoeyde - faktisk_takt, bottom=faktisk_takt, color='none', edgecolor='green', hatch='//')
            ax.set_title(f'{dag.strftime("%d.%m.%Y")}', fontsize=10)

        # Plotting weekly average data
        avg_stopptid = np.mean([data[1] for data in daglig_data])
        avg_arbeidstimer = np.mean([data[2] for data in daglig_data])
        avg_antall_fisk = np.mean([data[3] for data in daglig_data])
        avg_stopptid_impact = avg_stopptid * oee_100
        avg_stopptid_takt = round(avg_stopptid_impact / 60 / 8, 2)
        avg_faktisk_takt = round(avg_antall_fisk / avg_arbeidstimer, 2)
        avg_kjente_faktorer = round(avg_stopptid_takt, 2)
        avg_annet = oee_100 - avg_kjente_faktorer - avg_faktisk_takt
        avg_annet = round(avg_annet, 2)
        
        ax = axes.flatten()[-1]
        ax.bar(['100% OEE', 'Stopptid', 'Annet'], [oee_100, -avg_stopptid_takt, -avg_annet], color=['blue', 'red', 'orange'])
        ax.bar('Takttid', avg_faktisk_takt, bottom=0, color='green', edgecolor='black')
        ax.bar('Takttid', stiplet_hoeyde - avg_faktisk_takt, bottom=avg_faktisk_takt, color='none', edgecolor='green', hatch='//')
        ax.set_title(f'Ukesnitt {år}-W{uke_nummer}', fontsize=10)

        plt.tight_layout()
        st.pyplot(fig)

if __name__ == "__main__":
    main()
