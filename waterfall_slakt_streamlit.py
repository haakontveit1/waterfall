import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import matplotlib.pyplot as plt
import streamlit as st
from openpyxl import load_workbook

#fikse løsning dersom tidspunkt for slutt ikke er i riktig format



def les_data(uploaded_file):
    if uploaded_file is not None:
        try:
            # Load the data into a DataFrame (standard approach)
            df = pd.read_excel(uploaded_file, header=2)  # Skip 2 header rows
            
            # Reload the file using openpyxl to extract comments
            workbook = load_workbook(uploaded_file)
            sheet = workbook.active

            # Extract comments from the 4th column
            comments = []
            for idx, row in enumerate(sheet.iter_rows(min_row=3, min_col=4, max_col=4), start=0):
                cell = row[0]
                if cell.comment:
                    comment_text = cell.comment.text
                    # Extract the part of the comment after the third colon
                    colon_count = 0
                    extracted_text = ""
                    for char in comment_text:
                        if char == ":":
                            colon_count += 1
                        if colon_count == 3 or colon_count == 4:
                            extracted_text += char
                    if extracted_text:
                        comment_text = extracted_text.strip(": ")
                    else:
                        comment_text = ""
                else:
                    comment_text = ""  # No comment
                
                comments.append(comment_text)

            # Apply the workaround: Remove the first comment to align with DataFrame
            aligned_comments = comments[1:] + [""]  # Shift comments to align with DataFrame rows

            # Add aligned comments to the DataFrame
            df['comments'] = aligned_comments[:len(df)]  # Ensure exact match in length
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
                row.iloc[32:52].fillna(0).sum()
            )
        st.write(stopptid)
        return stopptid
    except Exception as e:
        st.error(f"Feil ved beregning av stopptid: {e}")
        return None

def beregn_faktiskproduksjon(row, sheet_type):
    try:
        # Parse start time
        start_time = datetime.strptime(str(row.iloc[2]), "%H:%M:%S")
        end_time_cell_value = str(row.iloc[3])
        st.write(f"Start time: {start_time}, End time (raw): {end_time_cell_value}")

        # Initialize work duration and parsed time
        work_duration = None
        parsed_time = None

        # Handle edge cases for "23:59:00" or "00:00:00"
        if end_time_cell_value in ["23:59:00", "00:00:00"]:
            st.write(f"Edge case detected for end time: {end_time_cell_value}")

            # Check for a comment in the 'comments' column
            comment_text = row['comments']
            if comment_text:
                st.write(f"Comment found: {comment_text}")

                # Extract hours and minutes from the comment
                try:
                    hh, mm = "", ""
                    for char in comment_text:
                        if char.isdigit() and len(hh) < 2:
                            hh += char
                        elif char.isdigit() and len(mm) < 2:
                            mm += char
                    if not hh or not mm:
                        raise ValueError("Failed to extract hh:mm from the comment")

                    # Convert to integers
                    hh = int(hh)
                    mm = int(mm)
                    parsed_time = timedelta(hours=hh, minutes=mm)
                    st.write(f"Parsed time from comment: {hh:02}:{mm:02} (timedelta: {parsed_time})")
                except Exception as e:
                    st.error(f"Could not parse time from comment: {e}")
                    return None, None

            # Calculate work duration using the parsed time
            if parsed_time:
                if end_time_cell_value == "23:59:00":
                    work_duration = timedelta(minutes=1) + parsed_time + timedelta(hours=23, minutes=59)
                    #st.write(f"23:59 Case with comment: work_duration = {work_duration}")
                elif end_time_cell_value == "00:00:00":
                    work_duration = parsed_time + timedelta(hours=24)
                    #st.write(f"00:00 Case with comment: work_duration = {work_duration}")
            else:
                # Fallback for no comment
                if end_time_cell_value == "23:59:00":
                    work_duration = timedelta(hours=23, minutes=59)
                    #st.write(f"No comment - 23:59 Case: work_duration = {work_duration}")
                elif end_time_cell_value == "00:00:00":
                    work_duration = timedelta(hours=24)
                    #st.write(f"No comment - 00:00 Case: work_duration = {work_duration}")
                    
            # Adjust final work duration to account for the start time
            work_duration = work_duration - (start_time - datetime.combine(start_time.date(), datetime.min.time()))
            #st.write(f"Final adjusted work_duration (after subtracting start_time): {work_duration}")
            #st.write(f"And the start time of the day was {start_time}")
            #st.write(f"And the start time of the day was {datetime.combine(start_time.date(), datetime.min.time())}")
        else:
            # Normal case for end time
            end_time = datetime.strptime(end_time_cell_value, "%H:%M:%S")
            work_duration = end_time - start_time
            #st.write(f"Normal case: work_duration = end_time - start_time = {work_duration}")

        # Convert work duration to minutes
        arbeidstimer = work_duration.total_seconds() / 60
        #st.write(f"arbeidstimer = work_duration.total_seconds() / 60 = {arbeidstimer} minutes")

        # Get fish count
        antall_fisk = row.iloc[4] if sheet_type == "slakt" else row.iloc[12]
        #st.write(f"antall_fisk = {antall_fisk}")

        # Return results
        return arbeidstimer, antall_fisk

    except Exception as e:
        st.error(f"Error in production time calculation: {e}")
        return None, None



def velg_dato():
    år = st.number_input("Velg år:", min_value=2024, max_value=datetime.now().year)
    maneder = ["Januar", "Februar", "Mars", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Desember"]
    maaned = st.selectbox("Velg måned:", list(range(1, 13)), format_func=lambda x: maneder[x-1])
    dag = st.number_input("Velg dag (1-31):", min_value=1, max_value=31)
    valgt_dato = datetime(år, maaned, dag)
    return valgt_dato

def hent_uke_dager(år, uke_nummer):
    # Find the Thursday of the previous week
    torsdag_forrige_uke = datetime.strptime(f'{år}-W{uke_nummer-1}-4', "%Y-W%W-%w")
    
    # Collect days: Thursday and Friday from the previous week, and Monday to Wednesday from the current week
    dager = [torsdag_forrige_uke + timedelta(days=i) for i in range(2)]  # Thursday and Friday
    mandag_naavaerende_uke = datetime.strptime(f'{år}-W{uke_nummer}-1', "%Y-W%W-%w")
    dager += [mandag_naavaerende_uke + timedelta(days=i) for i in range(3)]  # Monday, Tuesday, Wednesday
    
    return dager


def main():
    st.title("Produksjonsanalyse")

    fisk = "fisk"
    pa = " (på slakt)"

    # Velg type ark
    sheet_type = st.selectbox("Velg type ark:", ["slakt", "filet"])
    if sheet_type == "filet":
        fisk = "fileter"
        pa = " (på filet)"

    uploaded_file = st.file_uploader(f"Velg en Excel-fil (må være et 'input-{sheet_type}'-ark).", type=["xlsx"])
    
    # Velg type analyse
    analysis_type = st.selectbox("Velg analyse:", ["Spesifikk dato", "Hele uken"])
    oee_100 = 150 if sheet_type == "slakt" else 25
    stiplet_hoeyde = 120 if sheet_type == "slakt" else 20
    
    if analysis_type == "Spesifikk dato":
        valgt_dato = velg_dato()
    else:
        year = st.number_input("Velg år:", min_value=2024, max_value=datetime.now().year)
        week_number = st.number_input("Velg uke nummer:", min_value=1, max_value=52)
        week_days = hent_uke_dager(year, week_number)
    
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

    if analysis_type == "Spesifikk dato":
        valgt_dato_enkel = valgt_dato.date()
        if valgt_dato_enkel in df['Dato'].values:
            row = df[df['Dato'] == valgt_dato_enkel].iloc[0]
            stopptid = beregn_stopptid(row, sheet_type)
            arbeidstimer, antall_fisk = beregn_faktiskproduksjon(row, sheet_type)
            if stopptid is None or arbeidstimer is None or antall_fisk is None:
                st.error("Kan ikke beregne verdier. Sjekk om du har valgt riktig filtype og lastet opp riktig fil.")
                return
            stopptid_impact = stopptid * oee_100
            stopptid_takt = round(stopptid_impact / arbeidstimer, 2)
            faktisk_takt = round(antall_fisk / arbeidstimer, 2)
            kjente_faktorer = round(stopptid_takt, 2)
            annet = oee_100 - kjente_faktorer - faktisk_takt
            annet = round(annet, 2)
            
            # Plotting single day data
            stages = ['100% OEE', 'Stopptid', 'Annet']
            values = [oee_100, -stopptid_takt, -annet]

            cum_values = np.cumsum([0] + values).tolist()
            value_starts = cum_values[:-1]

            fig, ax = plt.subplots(figsize=(10, 5), dpi=100)  # Juster figurstørrelse og DPI
            colors = ['blue', 'red', 'orange']

            for i in range(len(stages)):
                ax.bar(stages[i], values[i], bottom=value_starts[i], color=colors[i], edgecolor='black')

            ax.bar('Takttid', faktisk_takt, bottom=0, color='green', edgecolor='black')
            ax.bar('Takttid', stiplet_hoeyde - faktisk_takt, bottom=faktisk_takt, color='none', edgecolor='green', hatch='//')

            for i in range(len(stages)):
                percentage = abs(values[i]) / oee_100 * 100  # Calculate percentage
                y_pos = value_starts[i] + values[i] / 2
                ax.text(
                    stages[i], y_pos,
                    f'{values[i]} ({percentage:.1f}%)',  # Include units and percentage
                    ha='center', va='center', color='white', fontweight='bold'
                )
            
            ax.text(
                'Takttid', faktisk_takt / 2,
                f'{faktisk_takt} ({(faktisk_takt / oee_100 * 100):.1f}%)',  # Include units and percentage
                ha='center', va='center', color='white', fontweight='bold'
            )
            
            # Add gap value above the stippled bar
            gap_to_80 = stiplet_hoeyde - faktisk_takt
            ax.text(
                'Takttid', stiplet_hoeyde + 5,  # Position above stippled bar
                f'{round(gap_to_80, 2)} ({(gap_to_80 / oee_100 * 100):.1f}%)',  # Gap value and percentage
                ha='center', va='bottom', color='green', fontweight='bold'
            )

            ax.set_ylabel(f'Antall {fisk} produsert per minutt')
            ax.set_title(f'Antall {fisk} produsert per minutt {valgt_dato.strftime("%d.%m.%Y")}{pa}')
            st.pyplot(fig)

        else:
            st.warning("Datoen du valgte finnes ikke i input-arket. Dette er enten fordi du tastet inn en ugyldig dato eller fordi datoen ikke hadde noen produksjon (eks helg).")
    else:
        daglig_data = []
        for dag in week_days:
            dag_enkel = dag.date()
            if dag_enkel in df['Dato'].values:
                row = df[df['Dato'] == dag_enkel].iloc[0]
                stopptid = beregn_stopptid(row, sheet_type)
                arbeidstimer, antall_fisk = beregn_faktiskproduksjon(row, sheet_type)
                if stopptid is None or arbeidstimer is None or antall_fisk is None:
                    st.error(f"Kan ikke beregne verdier for {dag.strftime('%d.%m.%Y')}. Sjekk om du har valgt riktig filtype og lastet opp riktig fil.")
                    return
                daglig_data.append((dag, stopptid, arbeidstimer, antall_fisk))
                # Plot daily graph
                st.write(stopptid)
                stopptid_impact = stopptid * oee_100
                stopptid_takt = round(stopptid_impact / arbeidstimer, 2)
                faktisk_takt = round(antall_fisk / arbeidstimer, 2)
                kjente_faktorer = round(stopptid_takt, 2)
                annet = oee_100 - kjente_faktorer - faktisk_takt
                annet = round(annet, 2)

                fig, ax = plt.subplots(figsize=(10, 5), dpi=100)
                stages = ['100% OEE', 'Stopptid', 'Annet']
                values = [oee_100, -stopptid_takt, -annet]
                cum_values = np.cumsum([0] + values).tolist()
                value_starts = cum_values[:-1]
                colors = ['blue', 'red', 'orange']

                for i in range(len(stages)):
                    ax.bar(stages[i], values[i], bottom=value_starts[i], color=colors[i], edgecolor='black')

                ax.bar('Takttid', faktisk_takt, bottom=0, color='green', edgecolor='black')
                ax.bar('Takttid', stiplet_hoeyde - faktisk_takt, bottom=faktisk_takt, color='none', edgecolor='green', hatch='//')

                for i in range(len(stages)):
                    percentage = abs(values[i]) / oee_100 * 100
                    y_pos = value_starts[i] + values[i] / 2
                    ax.text(
                        stages[i], y_pos,
                        f'{values[i]} ({percentage:.1f}%)',
                        ha='center', va='center', color='white', fontweight='bold'
                    )

                ax.text(
                    'Takttid', faktisk_takt / 2,
                    f'{faktisk_takt} ({(faktisk_takt / oee_100 * 100):.1f}%)',
                    ha='center', va='center', color='white', fontweight='bold'
                )

                gap_to_80 = stiplet_hoeyde - faktisk_takt
                ax.text(
                    'Takttid', stiplet_hoeyde + 5,
                    f'{round(gap_to_80, 2)} ({(gap_to_80 / oee_100 * 100):.1f}%)',
                    ha='center', va='bottom', color='green', fontweight='bold'
                )

                ax.set_ylabel(f'Antall {fisk} produsert per minutt')
                ax.set_title(f'Daglig produksjon {dag.strftime("%d.%m.%Y")}')
                st.write(f"Totalt antall arbeidstimer: {arbeidstimer}")
                st.pyplot(fig)

        if not daglig_data:
            st.warning("Ingen gyldige data funnet for den valgte uken.")
            return

        # Weekly averages
        avg_stopptid = np.mean([data[1] for data in daglig_data])
        avg_arbeidstimer = np.mean([data[2] for data in daglig_data])
        avg_antall_fisk = np.mean([data[3] for data in daglig_data])

        avg_stopptid_impact = avg_stopptid * oee_100
        avg_stopptid_takt = round(avg_stopptid_impact / avg_arbeidstimer, 2)
        avg_faktisk_takt = round(avg_antall_fisk / avg_arbeidstimer, 2)
        avg_kjente_faktorer = round(avg_stopptid_takt, 2)
        avg_annet = oee_100 - avg_kjente_faktorer - avg_faktisk_takt
        avg_annet = round(avg_annet, 2)

        # Plot weekly average graph
        fig, ax = plt.subplots(figsize=(10, 5), dpi=100)
        stages = ['100% OEE', 'Stopptid', 'Annet']
        values = [oee_100, -avg_stopptid_takt, -avg_annet]
        cum_values = np.cumsum([0] + values).tolist()
        value_starts = cum_values[:-1]
        colors = ['blue', 'red', 'orange']
        
        for i in range(len(stages)):
            ax.bar(stages[i], values[i], bottom=value_starts[i], color=colors[i], edgecolor='black')

        ax.bar('Takttid', avg_faktisk_takt, bottom=0, color='green', edgecolor='black')
        ax.bar('Takttid', stiplet_hoeyde - avg_faktisk_takt, bottom=avg_faktisk_takt, color='none', edgecolor='green', hatch='//')

        for i in range(len(stages)):
            percentage = abs(values[i]) / oee_100 * 100
            y_pos = value_starts[i] + values[i] / 2
            ax.text(
                stages[i], y_pos,
                f'{values[i]} ({percentage:.1f}%)',
                ha='center', va='center', color='white', fontweight='bold'
            )

        ax.text(
            'Takttid', avg_faktisk_takt / 2,
            f'{avg_faktisk_takt} ({(avg_faktisk_takt / oee_100 * 100):.1f}%)',
            ha='center', va='center', color='white', fontweight='bold'
        )

        gap_to_80 = stiplet_hoeyde - avg_faktisk_takt
        ax.text(
            'Takttid', stiplet_hoeyde + 5,
            f'{round(gap_to_80, 2)} ({(gap_to_80 / oee_100 * 100):.1f}%)',
            ha='center', va='bottom', color='green', fontweight='bold'
        )

        ax.set_ylabel(f'Antall {fisk} produsert per minutt')
        ax.set_title(f'Ukesnitt {year}-W{week_number}')
        st.pyplot(fig)
        

if __name__ == "__main__":
    main()
