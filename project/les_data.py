def les_data(uploaded_file):
    if uploaded_file is not None:
        try:
            if sheet_type == "slakt":
                df = pd.read_excel(uploaded_file, header=2)
                workbook = load_workbook(uploaded_file)
                sheet = workbook.active
                comments = [] # Forklaring i: project/les_data_comments.md
                for idx, row in enumerate(sheet.iter_rows(min_row=3, min_col=4, max_col=4), start=0):
                    cell = row[0]
                    if cell.comment:
                        comment_text = cell.comment.text
                        colon_count = 0 # Forklaring i:
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
                df = pd.read_excel(uploaded_file, header=2)

            return df
        except Exception as e:
            st.error(f"Feil ved lesing av Excel-filen: {e}")
            return None
    return None
