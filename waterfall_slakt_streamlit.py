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
    uploaded_file = st.file_uploader(f"Velg en Excel-fil (må være et 'input-{sheet_type}'-ark).", type=["xlsx"])
    
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
    valgt_dato_str = valgt_dato.strftime("%d.%m.%Y")  # Konverterer til ønsket format
    st.write(f"Valgt dato: {valgt_dato_str}")
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
        value_width = 70  # Adjust this value based on your value lengths

        faktisk_takt = round(antall_fisk / arbeidstimer, 2)
        kjente_faktorer = round(stopptid_takt, 2)
        annet = oee_100 - kjente_faktorer - faktisk_takt
        annet = round(annet, 2)

        # Define the formatted output using HTML
        output_html = f"""
        <table style="width:100%; border-spacing: 0;">
        <tr>
            <td style="text-align:left; width:70%; padding-right: 10px;">OEE 100%:</td>
            <td style="text-align:right; width:30%;">{oee_100} fisk/minutt</td>
        </tr>
        <tr>
            <td style="text-align:left; width:70%; padding-right: 10px;">Total stopptid:</td>
            <td style="text-align:right; width:30%;">{round(stopptid, 2)} minutter</td>
        </tr>
        <tr>
            <td style="text-align:left; width:70%; padding-right: 10px;">Tapt takt per minutt på grunn av stopp:</td>
            <td style="text-align:right; width:30%;">{stopptid_takt} fisk/minutt</td>
        </tr>
        </table>
        <br>  <!-- Tomme linjer for å skille seksjonene -->
        <table style="width:100%; border-spacing: 0;">
        <tr>
            <td style="text-align:left; width:70%; padding-right: 10px;">Arbeitstimer:</td>
            <td style="text-align:right; width:30%;">{round(arbeidstimer/60, 2)} timer</td>
        </tr>
        <tr>
            <td style="text-align:left; width:70%; padding-right: 10px;">Antall fisk produsert:</td>
            <td style="text-align:right; width:30%;">{antall_fisk} fisk</td>
        </tr>
        <tr>
            <td style="text-align:left; width:70%; padding-right: 10px;">Antall fisk tapt pga stopptid:</td>
            <td style="text-align:right; width:30%;">{round(stopptid_impact, 2)} fisk</td>
        </tr>
        </table>
        <br>
        <table style="width:100%; border-spacing: 0;">
        <tr>
            <td style="text-align:left; width:70%; padding-right: 10px;">Annet tap (unoterte feil, operatørhastighet etc):</td>
            <td style="text-align:right; width:30%;">{annet} minutter</td>
        </tr>
        <tr>
            <td style="text-align:left; width:70%; padding-right: 10px;">Faktisk takt:</td>
            <td style="text-align:right; width:30%;">{faktisk_takt} fisk/minutt</td>
        </tr>
        <tr>
            <td style="text-align:left; width:70%; padding-right: 10px;">Avstand fra 80% OEE takttid:</td>
            <td style="text-align:right; width:30%;">{round(stiplet_hoeyde - faktisk_takt, 2)} fisk/minutt</td>
        </tr>
        </table>
        """

        # Display the output in Streamlit using st.markdown
        st.markdown(output_html, unsafe_allow_html=True)

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

        fisk = "fisk"
        pa = " (på slakt)"
        if sheet_type == "filet":
            fisk = "fileter"
            pa = ""

        ax.set_ylabel(f'Antall {fisk} produsert per minutt')
        ax.set_title(f'Antall {fisk} produsert per minutt {valgt_dato_str}{pa}')
        st.pyplot(fig)

    else:
        st.warning("Datoen du valgte finnes ikke i input-arket. Dette er enten fordi du tastet inn en ugyldig dato eller fordi datoen ikke hadde noen produksjon (eks helg).")

if __name__ == "__main__":
    main()
