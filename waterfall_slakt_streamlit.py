import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import matplotlib.pyplot as plt
import streamlit as st
from openpyxl import load_workbook


def les_data(uploaded_file, sheet_type):
    if uploaded_file is not None:
        try:
            if sheet_type == "slakt":
                # Load data and extract comments for "slakt"
                df = pd.read_excel(uploaded_file, header=2)
                workbook = load_workbook(uploaded_file)
                sheet = workbook.active

                comments = []
                for idx, row in enumerate(sheet.iter_rows(min_row=3, min_col=4, max_col=4), start=0):
                    cell = row[0]
                    if cell.comment:
                        comment_text = cell.comment.text
                        colon_count = 0
                        hh_mm = ""
                        for char in comment_text:
                            if char == ":":
                                colon_count += 1
                            if colon_count >= 3 and char.isdigit():
                                hh_mm += char
                        comment_text = hh_mm.strip() if hh_mm else ""
                    else:
                        comment_text = ""
                    comments.append(comment_text)

                aligned_comments = comments[1:] + [""]
                df["comments"] = aligned_comments[:len(df)]
            else:
                # Simple load for "filet"
                df = pd.read_excel(uploaded_file, header=2)

            return df
        except Exception as e:
            st.error(f"Feil ved lesing av Excel-filen: {e}")
            return None
    return None


def velg_dato():
    år = st.number_input("Velg år:", min_value=2024, max_value=datetime.now().year)
    maneder = ["Januar", "Februar", "Mars", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Desember"]
    maaned = st.selectbox("Velg måned:", list(range(1, 13)), format_func=lambda x: maneder[x-1])
    dag = st.number_input("Velg dag (1-31):", min_value=1, max_value=31)
    valgt_dato = datetime(år, maaned, dag)
    return valgt_dato


def hent_uke_dager(år, uke_nummer, df):
    dager = []

    if uke_nummer == 1:
        # Handle the edge case for week 1 by looking at the last week of the previous year
        forrige_år = år - 1
        uke_forrige_år = 52
        if datetime.strptime(f'{forrige_år}-12-28', "%Y-%m-%d").isocalendar()[1] == 53:
            uke_forrige_år = 53
        torsdag_forrige_uke = datetime.strptime(f'{forrige_år}-W{uke_forrige_år}-4', "%Y-W%W-%w")
    else:
        # Normal case: Find Thursday of the previous week
        torsdag_forrige_uke = datetime.strptime(f'{år}-W{uke_nummer-1}-4', "%Y-W%W-%w")

    # Add Thursday and Friday from the previous week
    dager += [torsdag_forrige_uke, torsdag_forrige_uke + timedelta(days=1)]  # Thursday and Friday

    # Check and add Saturday and Sunday from the previous week if they exist in the production data
    lørdag_forrige_uke = torsdag_forrige_uke + timedelta(days=2)  # Saturday
    søndag_forrige_uke = torsdag_forrige_uke + timedelta(days=3)  # Sunday
    if lørdag_forrige_uke.date() in df['Dato'].values:
        dager.append(lørdag_forrige_uke)
    if søndag_forrige_uke.date() in df['Dato'].values:
        dager.append(søndag_forrige_uke)

    # Add Monday, Tuesday, and Wednesday from the current week
    mandag_naavaerende_uke = datetime.strptime(f'{år}-W{uke_nummer}-1', "%Y-W%W-%w")
    dager += [mandag_naavaerende_uke + timedelta(days=i) for i in range(3)]  # Monday, Tuesday, Wednesday

    return dager

def hent_maned_dager(år, maned, df):
    dager = []
    try:
        for _, row in df.iterrows():
            dato = row['Dato']
            if dato.year == år and dato.month == maned:
                dager.append(dato)
    except Exception as e:
        st.error(f"Feil ved henting av dager for måneden: {e}")
    return dager



def beregn_stopptid(row, sheet_type):
    try:
        if sheet_type == "slakt":
            stopptid = (
                row.iloc[27:31].fillna(0).sum() +
                row.iloc[34:40].fillna(0).sum() / 6 +
                row.iloc[40:51].fillna(0).sum()
            )
        elif sheet_type == "filet":
            stopptid = row.iloc[32:52].fillna(0).sum()
        return stopptid
    except Exception as e:
        st.error(f"Feil ved beregning av stopptid: {e}")
        return None

# Alt funker ish
# def beregn_faktiskproduksjon(row, sheet_type):
#     try:
#         if sheet_type == "slakt":
#             # Handle edge cases for "slakt"
#             start_time = datetime.strptime(str(row.iloc[2]), "%H:%M:%S")
#             end_time_cell_value = str(row.iloc[3])

#             if end_time_cell_value in ["23:59:00", "00:00:00"]:
#                 st.write(f"Edge case detected for end time: {end_time_cell_value}")
#                 comment_text = row['comments']

#                 if comment_text:
#                     st.write(f"Comment found: {comment_text}")
#                     # Extract hh:mm from the comment
#                     try:
#                         hh = int(comment_text[:2])
#                         mm = int(comment_text[3:5])
#                         parsed_time = timedelta(hours=hh, minutes=mm)
#                         st.write(f"Parsed time from comment: {parsed_time}")
#                     except Exception as e:
#                         st.error(f"Could not parse time from comment: {e}")
#                         return None, None

#                     if end_time_cell_value == "23:59:00":
#                         end_time = timedelta(minutes=1) + parsed_time + timedelta(hours=23, minutes=59)
#                     else:  # "00:00:00"
#                         end_time = parsed_time + timedelta(hours=24)

#                     work_duration = end_time - timedelta(hours=start_time.hour, minutes=start_time.minute, seconds=start_time.second)
#                 else:
#                     # No comment; fallback to default behavior
#                     end_time = datetime.strptime(end_time_cell_value, "%H:%M:%S")
#                     work_duration = end_time - start_time
#             else:
#                 # Normal case
#                 end_time = datetime.strptime(end_time_cell_value, "%H:%M:%S")
#                 work_duration = end_time - start_time

#             arbeidstimer = work_duration.total_seconds() / 60
#             antall_fisk = row.iloc[4]

#         elif sheet_type == "filet":
#             # Simpler case for "filet"
#             start_time = datetime.strptime(str(row.iloc[6]), "%H:%M:%S")
#             end_time = datetime.strptime(str(row.iloc[7]), "%H:%M:%S")
#             work_duration = end_time - start_time
#             arbeidstimer = work_duration.total_seconds() / 60
#             antall_fisk = row.iloc[12]

#         # Debug: Output calculated values
#         #st.write(f"Oppstartstid: {start_time.strftime('%H:%M:%S')}, Siste kasse på rework: {end_time.strftime('%H:%M:%S')}")
#         #st.write(f"Work duration: {work_duration}")
#         #st.write(f"Arbeidstimer (minutes): {arbeidstimer}")
#         #st.write(f"Antall fisk: {antall_fisk}")

#         return arbeidstimer, antall_fisk
#     except Exception as e:
#         st.error(f"Feil ved beregning av faktisk produksjon: {e}")
#         return None, None

# Edge cases funker:
# def beregn_faktiskproduksjon(row, sheet_type):
#     try:
#         if sheet_type == "slakt":
#             # Handle edge cases for "slakt"
#             start_time = datetime.strptime(str(row.iloc[2]), "%H:%M:%S")
#             end_time_cell_value = str(row.iloc[3])

#             if end_time_cell_value in ["23:59:00", "00:00:00"]:
#                 st.write(f"Edge case detected for end time: {end_time_cell_value}")
#                 comment_text = row['comments']

#                 if comment_text:
#                     st.write(f"Comment found: {comment_text}")
#                     # Extract hours and minutes from the comment
#                     try:
#                         if len(comment_text) == 4 and comment_text.isdigit():
#                             # If the comment is in "hhmm" format, extract hours and minutes
#                             hh = int(comment_text[:2])
#                             mm = int(comment_text[2:])
#                             parsed_time = timedelta(hours=hh, minutes=mm)
#                         elif ":" in comment_text:
#                             # If the comment is in "hh:mm" format, split and parse
#                             hh, mm = map(int, comment_text.split(":"))
#                             parsed_time = timedelta(hours=hh, minutes=mm)
#                         else:
#                             raise ValueError("Invalid time format in comment.")
#                         st.write(f"Parsed time from comment: {parsed_time}")
#                     except Exception as e:
#                         st.error(f"Could not parse time from comment: {e}")
#                         return None, None

#                     # Handle edge cases based on end_time_cell_value
#                     if end_time_cell_value == "23:59:00":
#                         end_time = timedelta(minutes=1) + parsed_time
#                     else:  # "00:00:00"
#                         end_time = timedelta(hours=24) + parsed_time

#                     # Adjust end_time to be one day ahead of start_time
#                     end_time_datetime = datetime.combine(start_time.date(), datetime.min.time()) + end_time
#                     if end_time_datetime < start_time:
#                         end_time_datetime += timedelta(days=1)

#                     # Calculate work duration
#                     work_duration = end_time_datetime - start_time
#                 else:
#                     # No comment; fallback to default behavior
#                     end_time = datetime.strptime(end_time_cell_value, "%H:%M:%S")
#                     work_duration = end_time - start_time
#             else:
#                 # Normal case
#                 end_time = datetime.strptime(end_time_cell_value, "%H:%M:%S")
#                 work_duration = end_time - start_time

#             arbeidstimer = work_duration.total_seconds() / 60
#             antall_fisk = row.iloc[4]

#         elif sheet_type == "filet":
#             # Simpler case for "filet"
#             start_time = datetime.strptime(str(row.iloc[6]), "%H:%M:%S")
#             end_time = datetime.strptime(str(row.iloc[7]), "%H:%M:%S")
#             work_duration = end_time - start_time
#             arbeidstimer = work_duration.total_seconds() / 60
#             antall_fisk = row.iloc[12]

#         # Debug: Output calculated values
#         st.write(f"Oppstartstid: {start_time.strftime('%H:%M:%S')}, Siste kasse på rework: {end_time_datetime.strftime('%H:%M:%S')}")
#         st.write(f"Work duration: {work_duration}")
#         st.write(f"Arbeidstimer (minutes): {arbeidstimer}")
#         st.write(f"Antall fisk: {antall_fisk}")

#         return arbeidstimer, antall_fisk
#     except Exception as e:
#         st.error(f"Feil ved beregning av faktisk produksjon: {e}")
#         return None, None

# Everything works
def beregn_faktiskproduksjon(row, sheet_type):
    try:
        if sheet_type == "slakt":
            # Handle edge cases for "slakt"
            start_time = datetime.strptime(str(row.iloc[2]), "%H:%M:%S")
            end_time_cell_value = str(row.iloc[3])

            if end_time_cell_value in ["23:59:00", "00:00:00"]:
                #st.write(f"Sluttidspunkt notert som: {end_time_cell_value}")
                comment_text = row['comments']

                if comment_text:
                    #st.write(f"Kommentar på sluttidspunkt-celle: {comment_text}")
                    # Extract hours and minutes from the comment
                    try:
                        if len(comment_text) == 4 and comment_text.isdigit():
                            # If the comment is in "hhmm" format, extract hours and minutes
                            hh = int(comment_text[:2])
                            mm = int(comment_text[2:])
                            parsed_time = timedelta(hours=hh, minutes=mm)
                        elif ":" in comment_text:
                            # If the comment is in "hh:mm" format, split and parse
                            hh, mm = map(int, comment_text.split(":"))
                            parsed_time = timedelta(hours=hh, minutes=mm)
                        else:
                            raise ValueError("Invalid time format in comment.")
                        #st.write(f"Sluttidspunkt trukket ut av kommentar: {parsed_time}")
                    except Exception as e:
                        st.error(f"Could not parse time from comment: {e}")
                        return None, None

                    # Handle edge cases based on end_time_cell_value
                    if end_time_cell_value == "23:59:00":
                        end_time = timedelta(minutes=0) + parsed_time
                    else:  # "00:00:00"
                        end_time = timedelta(hours=24) + parsed_time

                    # Adjust end_time to be one day ahead of start_time
                    end_time_datetime = datetime.combine(start_time.date(), datetime.min.time()) + end_time
                    if end_time_datetime < start_time:
                        end_time_datetime += timedelta(days=1)

                    # Calculate work duration
                    work_duration = end_time_datetime - start_time
                else:
                    st.error("Sluttidspunkt skrevet kan indikere sluttid etter kl 00:00, men ingen kommentar funnet.")
                    return None, None
            else:
                # Normal case
                end_time_datetime = datetime.strptime(end_time_cell_value, "%H:%M:%S")
                work_duration = end_time_datetime - start_time

            arbeidstimer = work_duration.total_seconds() / 60
            antall_fisk = row.iloc[4]

        elif sheet_type == "filet":
            # Simpler case for "filet"
            start_time = datetime.strptime(str(row.iloc[6]), "%H:%M:%S")
            end_time_datetime = datetime.strptime(str(row.iloc[7]), "%H:%M:%S")
            work_duration = end_time_datetime - start_time
            arbeidstimer = work_duration.total_seconds() / 60
            antall_fisk = row.iloc[12]

        # # Debug: Output calculated values
        # st.write(f"Oppstartstid: {start_time.strftime('%H:%M:%S')}, Siste kasse på rework: {end_time_datetime.strftime('%H:%M:%S')}")
        # st.write(f"Total arbeidstid: {work_duration}")
        # st.write(f"Antall fisk: {antall_fisk}")

        return arbeidstimer, antall_fisk
    except Exception as e:
        st.error(f"Feil ved beregning av faktisk produksjon: {e}")
        return None, None


def alle_grafer(month_days,df,sheet_type,oee_100,stiplet_hoeyde,fisk,selected_month,year):
    # Process daily data
    daglig_data = []
    for dag in month_days:
        row = df[df['Dato'] == dag].iloc[0]

        # Format with month as text
        formatted_date = dag.strftime("%d. %B %Y")
        
        # Print in Norwegian style
        st.write(f"Dato: {formatted_date}")
        stopptid = beregn_stopptid(row, sheet_type)
        arbeidstimer, antall_fisk = beregn_faktiskproduksjon(row, sheet_type)
        if stopptid is not None and arbeidstimer is not None and antall_fisk is not None:
            daglig_data.append((dag, stopptid, arbeidstimer, antall_fisk))

            # Plot daily graph
            st.write(f"Total stopptid i minutter: {round(stopptid,2)}")
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
                if stages[i] == 'Stopptid':
                    if sheet_type == "slakt":
                        if stopptid_takt < 9:
                            # Place the text outside the bar if the value is less than 7
                                dynamic_offset = value_starts[i] + values[i] - 7  # Adjust `-10` for spacing
                                ax.text(
                                    stages[i], dynamic_offset,
                                    f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                                    ha='center', va='top', color='black', fontweight='bold'
                                )
                        else:
                            # Place the text inside the bar if the value is greater than or equal to 7
                            y_pos = value_starts[i] + values[i] / 2
                            ax.text(
                                stages[i], y_pos,
                                f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                                ha='center', va='center', color='white', fontweight='bold'
                            )
                    elif sheet_type == "filet":
                        if stopptid_takt < 1.5:
                            # Place the text outside the bar if the value is less than 7
                            dynamic_offset = value_starts[i] + values[i] -1  # Adjust `-10` for spacing
                            ax.text(
                                stages[i], dynamic_offset,
                                f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                                ha='center', va='top', color='black', fontweight='bold'
                            )
                        else:
                            # Place the text inside the bar if the value is greater than or equal to 7
                            y_pos = value_starts[i] + values[i] / 2
                            ax.text(
                                stages[i], y_pos,
                                f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                                ha='center', va='center', color='white', fontweight='bold'
                            )
                else:
                    # Place the text inside other bars as before
                    y_pos = value_starts[i] + values[i] / 2
                    ax.text(
                        stages[i], y_pos,
                        f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
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
            st.pyplot(fig)


    # Print separator for monthly average
    st.write("---")
    st.header(f"Oppsummering for {selected_month} {year}")

    # Calculate monthly averages
    avg_stopptid = np.mean([data[1] for data in daglig_data])
    avg_arbeidstimer = np.mean([data[2] for data in daglig_data])
    avg_antall_fisk = np.mean([data[3] for data in daglig_data])

    avg_stopptid_impact = avg_stopptid * oee_100
    avg_stopptid_takt = round(avg_stopptid_impact / avg_arbeidstimer, 2)
    avg_faktisk_takt = round(avg_antall_fisk / avg_arbeidstimer, 2)
    avg_kjente_faktorer = round(avg_stopptid_takt, 2)
    avg_annet = oee_100 - avg_kjente_faktorer - avg_faktisk_takt
    avg_annet = round(avg_annet, 2)

    # Plot monthly average graph
    fig, ax = plt.subplots(figsize=(10, 5), dpi=100)
    stages = ['100% OEE', 'Stopptid', 'Annet']
    values = [oee_100, -avg_stopptid_takt, -avg_annet]
    cum_values = np.cumsum([0] + values).tolist()
    value_starts = cum_values[:-1]
    colors = ['blue', 'red', 'orange']
    
    for i in range(len(stages)):
        ax.bar(stages[i], values[i], bottom=value_starts[i], color=colors[i], edgecolor='black')
        y_pos = value_starts[i] + values[i] / 2
        if stages[i] == 'Stopptid' and avg_stopptid_takt < 9:
            dynamic_offset = value_starts[i] + values[i] - 7
            ax.text(
                stages[i], dynamic_offset,
                f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                ha='center', va='top', color='black', fontweight='bold'
            )
        else:
            ax.text(
                stages[i], y_pos,
                f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                ha='center', va='center', color='white', fontweight='bold'
            )
    
    # Add labels for Takttid
    ax.bar('Takttid', avg_faktisk_takt, bottom=0, color='green', edgecolor='black')
    ax.bar('Takttid', stiplet_hoeyde - avg_faktisk_takt, bottom=avg_faktisk_takt, color='none', edgecolor='green', hatch='//')
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
    ax.set_title(f'Månedsnitt for {selected_month} {year}')
    st.pyplot(fig)
    return

def manedlig(month_days,df,sheet_type,oee_100,stiplet_hoeyde,fisk,selected_month,year):
    # Process daily data
    daglig_data = []
    for dag in month_days:
        row = df[df['Dato'] == dag].iloc[0]
        stopptid = beregn_stopptid(row, sheet_type)
        arbeidstimer, antall_fisk = beregn_faktiskproduksjon(row, sheet_type)
        if stopptid is not None and arbeidstimer is not None and antall_fisk is not None:
            daglig_data.append((dag, stopptid, arbeidstimer, antall_fisk))
            
            


    # Print separator for monthly average
    st.write("---")
    st.header(f"Oppsummering for {selected_month} {year}")

    # Calculate monthly averages
    avg_stopptid = np.mean([data[1] for data in daglig_data])
    avg_arbeidstimer = np.mean([data[2] for data in daglig_data])
    avg_antall_fisk = np.mean([data[3] for data in daglig_data])

    avg_stopptid_impact = avg_stopptid * oee_100
    avg_stopptid_takt = round(avg_stopptid_impact / avg_arbeidstimer, 2)
    avg_faktisk_takt = round(avg_antall_fisk / avg_arbeidstimer, 2)
    avg_kjente_faktorer = round(avg_stopptid_takt, 2)
    avg_annet = oee_100 - avg_kjente_faktorer - avg_faktisk_takt
    avg_annet = round(avg_annet, 2)

    # Plot monthly average graph
    fig, ax = plt.subplots(figsize=(10, 5), dpi=100)
    stages = ['100% OEE', 'Stopptid', 'Annet']
    values = [oee_100, -avg_stopptid_takt, -avg_annet]
    cum_values = np.cumsum([0] + values).tolist()
    value_starts = cum_values[:-1]
    colors = ['blue', 'red', 'orange']
    
    for i in range(len(stages)):
        ax.bar(stages[i], values[i], bottom=value_starts[i], color=colors[i], edgecolor='black')
        y_pos = value_starts[i] + values[i] / 2
        if stages[i] == 'Stopptid' and avg_stopptid_takt < 9:
            dynamic_offset = value_starts[i] + values[i] - 7
            ax.text(
                stages[i], dynamic_offset,
                f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                ha='center', va='top', color='black', fontweight='bold'
            )
        else:
            ax.text(
                stages[i], y_pos,
                f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                ha='center', va='center', color='white', fontweight='bold'
            )
    
    # Add labels for Takttid
    ax.bar('Takttid', avg_faktisk_takt, bottom=0, color='green', edgecolor='black')
    ax.bar('Takttid', stiplet_hoeyde - avg_faktisk_takt, bottom=avg_faktisk_takt, color='none', edgecolor='green', hatch='//')
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
    ax.set_title(f'Månedsnitt for {selected_month} {year}')
    st.pyplot(fig)
    return



def main():
    st.title("Produksjonsanalyse")

    fisk = "fisk"
    pa = " (på slakt)"
    sheet_type = st.selectbox("Velg type ark:", ["slakt", "filet"])
    if sheet_type == "filet":
        fisk = "fileter"
        pa = " (på filet)"

    uploaded_file = st.file_uploader(f"Velg en Excel-fil (må være et 'input-{sheet_type}'-ark).", type=["xlsx"])
    analysis_type = st.selectbox("Velg analyse:", ["Spesifikk dato", "Hele uken", "Månedsrapport"])
    oee_100 = 150 if sheet_type == "slakt" else 25
    stiplet_hoeyde = 120 if sheet_type == "slakt" else 20

    if uploaded_file is None:
        st.warning("Vennligst last opp en Excel-fil for å fortsette.")
        return

    # Load and process the data
    df = les_data(uploaded_file, sheet_type)
    if df is None:
        st.warning("Ingen data tilgjengelig i den opplastede filen. Vennligst last opp en gyldig Excel-fil.")
        return

    # Ensure 'Dato' column exists
    try:
        df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], format="%Y-%m-%d %H:%M:%S")
        df['Dato'] = df.iloc[:, 0].dt.date
    except ValueError as e:
        st.error(f"Feil ved behandling av datoer: {e}")
        return

    if analysis_type == "Spesifikk dato":
        valgt_dato = velg_dato()
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
                if stages[i] == 'Stopptid':
                    if stopptid_takt < 7:
                        # Place the text outside the bar if the value is less than 7
                        dynamic_offset = value_starts[i] + values[i] - 10  # Adjust `-10` for spacing
                        ax.text(
                            stages[i], dynamic_offset,
                            f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                            ha='center', va='top', color='black', fontweight='bold'
                        )
                    else:
                        # Place the text inside the bar if the value is greater than or equal to 7
                        y_pos = value_starts[i] + values[i] / 2
                        ax.text(
                            stages[i], y_pos,
                            f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                            ha='center', va='center', color='white', fontweight='bold'
                        )
                else:
                    # Place the text inside other bars as before
                    y_pos = value_starts[i] + values[i] / 2
                    ax.text(
                        stages[i], y_pos,
                        f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
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
    elif analysis_type == "Hele uken":
        year = st.number_input("Velg år:", min_value=2024, max_value=datetime.now().year)
        week_number = st.number_input("Velg uke nummer:", min_value=1, max_value=52)
        week_days = hent_uke_dager(year, week_number, df)
               
        daglig_data = []
        for dag in week_days:
            dag_enkel = dag.date()
            if dag_enkel in df['Dato'].values:
                row = df[df['Dato'] == dag_enkel].iloc[0]
                # Format with month as text
                formatted_date = dag.strftime("%d. %B %Y")
                # Print in Norwegian style
                st.write(f"Dato: {formatted_date}")
                stopptid = beregn_stopptid(row, sheet_type)
                arbeidstimer, antall_fisk = beregn_faktiskproduksjon(row, sheet_type)
                if stopptid is None or arbeidstimer is None or antall_fisk is None:
                    st.error(f"Kan ikke beregne verdier for {dag.strftime('%d.%m.%Y')}. Sjekk om du har valgt riktig filtype og lastet opp riktig fil.")
                    return
                daglig_data.append((dag, stopptid, arbeidstimer, antall_fisk))
                # Plot daily graph
                st.write(f"Total stopptid i minutter: {round(stopptid,2)}")
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
                    if stages[i] == 'Stopptid':
                        if sheet_type == "slakt":
                            if stopptid_takt < 9:
                                # Place the text outside the bar if the value is less than 7
                                    dynamic_offset = value_starts[i] + values[i] - 7  # Adjust `-10` for spacing
                                    ax.text(
                                        stages[i], dynamic_offset,
                                        f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                                        ha='center', va='top', color='black', fontweight='bold'
                                    )
                            else:
                                # Place the text inside the bar if the value is greater than or equal to 7
                                y_pos = value_starts[i] + values[i] / 2
                                ax.text(
                                    stages[i], y_pos,
                                    f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                                    ha='center', va='center', color='white', fontweight='bold'
                                )
                        elif sheet_type == "filet":
                            if stopptid_takt < 1.5:
                                # Place the text outside the bar if the value is less than 7
                                    dynamic_offset = value_starts[i] + values[i] -1  # Adjust `-10` for spacing
                                    ax.text(
                                        stages[i], dynamic_offset,
                                        f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                                        ha='center', va='top', color='black', fontweight='bold'
                                    )
                            else:
                                # Place the text inside the bar if the value is greater than or equal to 7
                                y_pos = value_starts[i] + values[i] / 2
                                ax.text(
                                    stages[i], y_pos,
                                    f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                                    ha='center', va='center', color='white', fontweight='bold'
                                )
                    else:
                        # Place the text inside other bars as before
                        y_pos = value_starts[i] + values[i] / 2
                        ax.text(
                            stages[i], y_pos,
                            f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
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
            if stages[i] == 'Stopptid':
                if sheet_type == "slakt":
                    if avg_stopptid_takt < 9:
                        # Place the text outside the bar if the value is less than 7
                            dynamic_offset = value_starts[i] + values[i] - 7  # Adjust `-10` for spacing
                            ax.text(
                                stages[i], dynamic_offset,
                                f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                                ha='center', va='top', color='black', fontweight='bold'
                            )
                    else:
                        # Place the text inside the bar if the value is greater than or equal to 7
                        y_pos = value_starts[i] + values[i] / 2
                        ax.text(
                            stages[i], y_pos,
                            f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                            ha='center', va='center', color='white', fontweight='bold'
                        )
                elif sheet_type == "filet":
                    if avg_stopptid_takt < 1.5:
                        # Place the text outside the bar if the value is less than 7
                            dynamic_offset = value_starts[i] + values[i] -1  # Adjust `-10` for spacing
                            ax.text(
                                stages[i], dynamic_offset,
                                f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                                ha='center', va='top', color='black', fontweight='bold'
                            )
                    else:
                        # Place the text inside the bar if the value is greater than or equal to 7
                        y_pos = value_starts[i] + values[i] / 2
                        ax.text(
                            stages[i], y_pos,
                            f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
                            ha='center', va='center', color='white', fontweight='bold'
                        )
            else:
                # Place the text inside other bars as before
                y_pos = value_starts[i] + values[i] / 2
                ax.text(
                    stages[i], y_pos,
                    f'{values[i]} ({abs(values[i]) / oee_100 * 100:.1f}%)',
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
        
    else: #Hele måneden
        year = st.number_input("Velg år:", min_value=2024, max_value=datetime.now().year)
        months = ["Velg måned", "Januar", "Februar", "Mars", "April", "Mai", "Juni", 
                  "Juli", "August", "September", "Oktober", "November", "Desember"]
        selected_month = st.selectbox("Velg måned:", months)
        
        # Check if a valid month is selected
        if selected_month == "Velg måned":
            st.warning("Vennligst velg en måned for å fortsette.")
            return
        else:
            # Convert the selected month to its corresponding number
            month_number = months.index(selected_month)
            
            # Get all production days for the selected month
            month_days = hent_maned_dager(year, month_number, df)
        
            if not month_days:
                st.warning(f"Ingen produksjonsdager funnet for {selected_month} {year}.")
            else:
                st.write(f"Fant {len(month_days)} produksjonsdager for {selected_month} {year}")
        
        nedtrekk = ["Velg alternativ","Alle grafene","Kun månedlig gjennomsnitt"]
        graf_valg = st.selectbox("Vil du ha alle grafene eller kun månedlig gjennomsnitt?",nedtrekk)
        
        if graf_valg == "Velg alternativ":
            st.warning("Vennligst velg et alternativ for å fortsette.")
            return
        
        elif graf_valg == "Alle grafene":
            alle_grafer(month_days,df,sheet_type,oee_100,stiplet_hoeyde,fisk,selected_month,year)
        
        else: #Kun månedlig gjennomsnitt
            manedlig(month_days,df,sheet_type,oee_100,stiplet_hoeyde,fisk,selected_month,year)
            
            
if __name__ == "__main__":
    main()
