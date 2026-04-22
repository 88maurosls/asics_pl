import streamlit as st
import pandas as pd
import io
import gspread
from google.oauth2.service_account import Credentials

# Funzione per formattare la colonna Taglia
def format_taglia(size_us):
    if pd.isna(size_us):
        return ""
    size_str = str(size_us).strip()
    if size_str.endswith(".0"):
        size_str = size_str[:-2]
    return size_str.replace(".5", "+")

# Funzione per duplicare le righe in base al valore di Qta
def expand_rows(df):
    df = df.copy()
    df["Qta"] = pd.to_numeric(df["Qta"], errors="coerce").fillna(0).astype(int)
    df = df[df["Qta"] > 0]

    expanded_df = df.loc[df.index.repeat(df["Qta"])].assign(Qta=1)
    expanded_df["Tot Costo"] = ""
    return expanded_df

# Funzione per caricare il file color.txt e restituire un dizionario di mapping
def load_colors_mapping(file_path):
    colors_mapping = {}
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                if ';' in line:
                    try:
                        key, value = line.split(';', 1)
                        colors_mapping[key.strip().upper()] = value.strip()
                    except ValueError:
                        st.warning(f"Errore nel parsing della riga: {line}. Ignorata.")
                elif line:
                    st.warning(f"Riga malformata nel file color.txt: {line}. Ignorata.")
    except FileNotFoundError:
        st.warning("File color.txt non trovato. La colonna 'Base Color' resterà vuota.")
    return colors_mapping

# Funzione per determinare il valore di "Base Color"
def get_base_color(color_code, colors_mapping):
    if pd.isna(color_code):
        return ""
    color_code = str(color_code).strip().upper()
    for key in colors_mapping:
        if color_code.startswith(key):
            return colors_mapping[key]
    return ""

# Funzione per elaborare ogni file caricato
def process_file(file, colors_mapping):
    df = pd.read_excel(
        file,
        sheet_name="Delivery Items",
        dtype={
            "Item Code": str,
            "Color Code": str,
            "EAN Code": str
        }
    )

    df.columns = df.columns.astype(str).str.strip()

    required_columns = [
        "Item Code",
        "Color Code",
        "US Size",
        "EAN Code",
        "Item Description",
        "Delivery qty."
    ]

    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Nel file {file.name} mancano queste colonne: {missing_columns}")
        st.stop()

    output_df = pd.DataFrame({
        "Articolo": df["Item Code"].astype(str).str.strip(),
        "Descrizione": df["Item Description"].fillna("").astype(str).str.strip(),
        "Categoria": "CALZATURE",
        "Subcategoria": "Sneakers",
        "Colore": df["Color Code"].astype(str).str.strip().str.zfill(3),
        "Base Color": df["Color Code"].apply(lambda x: get_base_color(x, colors_mapping)),
        "Made in": "",
        "Sigla Bimbo": "",
        "Costo": "",
        "Retail": "",
        "Taglia": df["US Size"].apply(format_taglia),
        "Barcode": df["EAN Code"].fillna("").astype(str).str.strip(),
        "EAN": "",
        "Qta": pd.to_numeric(df["Delivery qty."], errors="coerce").fillna(0).astype(int),
        "Tot Costo": "",
        "Materiale": "",
        "Spec. Materiale": "",
        "Misure": "",
        "Scala Taglie": "US",
        "Tacco": "",
        "Suola": "",
        "Carryover": "",
        "HS Code": ""
    })

    output_df = output_df[output_df["Qta"] > 0]

    expanded_df = expand_rows(output_df)

    return expanded_df

# Funzione per suddividere i dati in fogli di massimo 50 righe e aggiungere l'intestazione
def write_data_in_chunks(writer, df, stagione, data_inizio, data_fine, ricarico):
    num_chunks = len(df) // 50 + (1 if len(df) % 50 > 0 else 0)

    for i in range(num_chunks):
        chunk_df = df[i * 50:(i + 1) * 50]
        sheet_name = f"Foglio{i+1}"
        start_row = 9

        chunk_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)

        if sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
        else:
            raise ValueError(f"Il foglio {sheet_name} non è stato trovato!")

        worksheet.write('A1', 'STAGIONE:')
        worksheet.write('B1', stagione)
        worksheet.write('A2', 'TIPO:')
        worksheet.write('B2', 'ACCESSORI')
        worksheet.write('A3', 'DATA INIZIO:')
        worksheet.write('B3', data_inizio.strftime('%d/%m/%Y'))
        worksheet.write('A4', 'DATA FINE:')
        worksheet.write('B4', data_fine.strftime('%d/%m/%Y'))
        worksheet.write('A5', 'RICARICO:')
        worksheet.write('B5', ricarico)

        text_format = writer.book.add_format({'num_format': '@'})
        worksheet.set_column('L:L', 20, text_format)

        last_data_row = len(chunk_df) + start_row
        empty_row = last_data_row + 1
        worksheet.write(f'N{empty_row}', "")
        worksheet.write(f'O{empty_row}', "")

        total_row = empty_row + 2
        worksheet.write(f'N{total_row}', "")
        worksheet.write(f'O{total_row}', "")

# Funzione per connettersi a Google Sheets
def connect_to_gsheet():
    credentials = {
        "type": st.secrets["gsheet"]["type"],
        "project_id": st.secrets["gsheet"]["project_id"],
        "private_key_id": st.secrets["gsheet"]["private_key_id"],
        "private_key": st.secrets["gsheet"]["private_key"],
        "client_email": st.secrets["gsheet"]["client_email"],
        "client_id": st.secrets["gsheet"]["client_id"],
        "auth_uri": st.secrets["gsheet"]["auth_uri"],
        "token_uri": st.secrets["gsheet"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["gsheet"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["gsheet"]["client_x509_cert_url"]
    }

    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(credentials, scopes=scope)
    client = gspread.authorize(creds)
    return client

# Funzione per recuperare i dati "Articolo", "Colore" e "Gender" dal foglio "Gender"
def get_existing_gender(sheet_url):
    client = connect_to_gsheet()
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.worksheet("Gender")

    data = worksheet.get_all_values()
    gender_dict = {f"{row[0]}-{row[1]}": row[2] for row in data[1:] if len(row) >= 3}
    return gender_dict

# Funzione per scrivere o aggiornare dati su Google Sheets
def write_to_gsheet(data, sheet_url):
    client = connect_to_gsheet()
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.worksheet("Gender")

    existing_data = worksheet.get_all_values()
    existing_entries = {
        f"{row[0]}-{row[1]}": idx + 2
        for idx, row in enumerate(existing_data[1:])
        if len(row) >= 2
    }

    batch_updates = []

    for (articolo, colore, gender) in data:
        key = f"{articolo}-{colore}"

        if key in existing_entries:
            row_to_update = existing_entries[key]
            batch_updates.append({
                "range": f'C{row_to_update}',
                "values": [[gender]]
            })
        else:
            worksheet.append_row([articolo, colore, gender])

    if batch_updates:
        worksheet.batch_update(batch_updates)

    st.success("Dati aggiornati o aggiunti su Google Sheet.")

# Streamlit app
st.title('Asics Xmag Lineare')

# Campi di input per l'intestazione
stagione = st.text_input("Inserisci STAGIONE")
data_inizio = st.date_input("Inserisci DATA INIZIO")
data_fine = st.date_input("Inserisci DATA FINE")
ricarico = st.text_input("Inserisci RICARICO", value="")  # lasciato libero ma non usato nei calcoli

# Link utile
st.markdown('**[Scarica le Packing List da qui](https://b2b.asics.com/orders-overview/order-history)**')

# Carica il file color.txt dalla directory del progetto
colors_mapping = load_colors_mapping("color.txt")

# Upload file
uploaded_files = st.file_uploader(
    "Scegli i file Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files and stagione and data_inizio and data_fine:
    processed_dfs = []

    google_sheet_url = "https://docs.google.com/spreadsheets/d/1p84nF9Tq-1ZJgQSEJcgrePLvQyGQ3cjt_1IZP5qPs00/edit?usp=sharing"
    gender_dict = get_existing_gender(google_sheet_url)

    for uploaded_file in uploaded_files:
        processed_dfs.append(process_file(uploaded_file, colors_mapping))

    if processed_dfs:
        final_df = pd.concat(processed_dfs, ignore_index=True)
    else:
        final_df = pd.DataFrame()

    if final_df.empty:
        st.warning("Nessun dato valido trovato nei file caricati.")
    else:
        unique_combinations = final_df[["Articolo", "Colore"]].drop_duplicates()

        st.write("Anteprima Articolo-Colore:")

        selections = {}

        gender_options = ["Seleziona...", "UOMO", "DONNA", "UNISEX"]

        for index, row in unique_combinations.iterrows():
            articolo_colore = f"{row['Articolo']}-{row['Colore']}"
            preselected_gender = gender_dict.get(articolo_colore, "Seleziona...")

            flag = st.selectbox(
                f"{row['Articolo']}-{row['Colore']}",
                options=gender_options,
                key=f"{row['Articolo']}_{row['Colore']}_{index}",
                index=gender_options.index(preselected_gender) if preselected_gender in gender_options else 0
            )

            selections[(row['Articolo'], row['Colore'])] = flag

        if st.button("Elabora File"):
            if any(flag == "Seleziona..." for flag in selections.values()):
                st.error("Devi selezionare UOMO, DONNA o UNISEX per tutte le combinazioni!")
            else:
                gsheet_data = [
                    (row['Articolo'], row['Colore'], selections[(row['Articolo'], row['Colore'])])
                    for _, row in unique_combinations.iterrows()
                ]

                write_to_gsheet(gsheet_data, google_sheet_url)

                uomo_df = final_df[
                    final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'UOMO', axis=1)
                ]
                donna_df = final_df[
                    final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'DONNA', axis=1)
                ]
                unisex_df = final_df[
                    final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'UNISEX', axis=1)
                ]

                uomo_output = io.BytesIO()
                donna_output = io.BytesIO()
                unisex_output = io.BytesIO()

                if not uomo_df.empty:
                    with pd.ExcelWriter(uomo_output, engine='xlsxwriter') as writer_uomo:
                        write_data_in_chunks(writer_uomo, uomo_df, stagione, data_inizio, data_fine, ricarico)
                    st.download_button(
                        label="Download File UOMO",
                        data=uomo_output.getvalue(),
                        file_name="uomo_processed_file.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                if not donna_df.empty:
                    with pd.ExcelWriter(donna_output, engine='xlsxwriter') as writer_donna:
                        write_data_in_chunks(writer_donna, donna_df, stagione, data_inizio, data_fine, ricarico)
                    st.download_button(
                        label="Download File DONNA",
                        data=donna_output.getvalue(),
                        file_name="donna_processed_file.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                if not unisex_df.empty:
                    with pd.ExcelWriter(unisex_output, engine='xlsxwriter') as writer_unisex:
                        write_data_in_chunks(writer_unisex, unisex_df, stagione, data_inizio, data_fine, ricarico)
                    st.download_button(
                        label="Download File UNISEX",
                        data=unisex_output.getvalue(),
                        file_name="unisex_processed_file.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
