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

# Funzione per pulire il prezzo
def clean_price_value(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""
    s = str(value).strip().replace("€", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return ""

# Funzione per duplicare le righe in base al valore di Qta
def expand_rows(df):
    df = df.copy()
    df["Qta"] = pd.to_numeric(df["Qta"], errors="coerce").fillna(0).astype(int)
    df = df[df["Qta"] > 0]

    expanded_df = df.loc[df.index.repeat(df["Qta"])].assign(Qta=1)
    expanded_df["Tot Costo"] = expanded_df["Costo"]
    return expanded_df

# Connessione a Google Sheets
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

# Recupera dati esistenti dal foglio Gender
def get_existing_data(sheet_url):
    client = connect_to_gsheet()
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.worksheet("Gender")

    data = worksheet.get_all_values()

    result = {}
    for row in data[1:]:
        articolo = row[0].strip() if len(row) > 0 else ""
        colore = row[1].strip() if len(row) > 1 else ""
        gender = row[2].strip() if len(row) > 2 else ""
        base_color = row[3].strip() if len(row) > 3 else ""
        price = row[4].strip() if len(row) > 4 else ""

        if articolo and colore:
            result[f"{articolo}-{colore}"] = {
                "gender": gender,
                "base_color": base_color,
                "price": price
            }

    return result

# Scrive o aggiorna dati su Google Sheets
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

    for articolo, colore, gender, base_color, price in data:
        key = f"{articolo}-{colore}"

        if key in existing_entries:
            row_to_update = existing_entries[key]
            batch_updates.append({
                "range": f"C{row_to_update}:E{row_to_update}",
                "values": [[gender, base_color, price]]
            })
        else:
            worksheet.append_row([articolo, colore, gender, base_color, price])

    if batch_updates:
        worksheet.batch_update(batch_updates)

    st.success("Dati aggiornati o aggiunti su Google Sheet.")

# Elabora file Excel
def process_file(file, memory_dict):
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

    # Mappa Base Color e Price da memoria Google Sheet
    keys = (
        df["Item Code"].astype(str).str.strip() + "-" +
        df["Color Code"].astype(str).str.strip().str.zfill(3)
    )

    base_colors = []
    prices = []

    for key in keys:
        row_data = memory_dict.get(key, {})
        base_colors.append(row_data.get("base_color", ""))
        prices.append(clean_price_value(row_data.get("price", "")))

    output_df = pd.DataFrame({
        "Articolo": df["Item Code"].astype(str).str.strip(),
        "Descrizione": df["Item Description"].fillna("").astype(str).str.strip(),
        "Categoria": "CALZATURE",
        "Subcategoria": "Sneakers",
        "Colore": df["Color Code"].astype(str).str.strip().str.zfill(3),
        "Base Color": base_colors,
        "Made in": "",
        "Sigla Bimbo": "",
        "Costo": prices,
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

# Scrittura Excel
def write_data_in_chunks(writer, df, stagione, data_inizio, data_fine, ricarico):
    num_chunks = len(df) // 50 + (1 if len(df) % 50 > 0 else 0)

    for i in range(num_chunks):
        chunk_df = df[i * 50:(i + 1) * 50]
        sheet_name = f"Foglio{i+1}"
        start_row = 9

        chunk_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)

        worksheet = writer.sheets[sheet_name]

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
        number_format = writer.book.add_format({'num_format': '#,##0.00'})

        worksheet.set_column('L:L', 20, text_format)
        worksheet.set_column('I:I', 12, number_format)
        worksheet.set_column('N:N', 12, number_format)

        last_data_row = len(chunk_df) + start_row
        empty_row = last_data_row + 1
        worksheet.write(f'N{empty_row}', "")
        worksheet.write(f'O{empty_row}', "")

        total_row = empty_row + 2
        worksheet.write_formula(f'N{total_row}', f"=SUM(N{start_row+2}:N{last_data_row + 1})", number_format)
        worksheet.write(f'O{total_row}', "")

# Streamlit app
st.title('Asics Xmag Lineare')

stagione = st.text_input("Inserisci STAGIONE")
data_inizio = st.date_input("Inserisci DATA INIZIO")
data_fine = st.date_input("Inserisci DATA FINE")
ricarico = st.text_input("Inserisci RICARICO", value="")

st.markdown('**[Scarica le Packing List da qui](https://b2b.asics.com/orders-overview/order-history)**')

uploaded_files = st.file_uploader(
    "Scegli i file Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files and stagione and data_inizio and data_fine:
    google_sheet_url = "https://docs.google.com/spreadsheets/d/1p84nF9Tq-1ZJgQSEJcgrePLvQyGQ3cjt_1IZP5qPs00/edit?usp=sharing"
    memory_dict = get_existing_data(google_sheet_url)

    preview_dfs = []
    for uploaded_file in uploaded_files:
        df_preview = pd.read_excel(
            uploaded_file,
            sheet_name="Delivery Items",
            dtype={"Item Code": str, "Color Code": str}
        )
        df_preview.columns = df_preview.columns.astype(str).str.strip()

        temp_df = pd.DataFrame({
            "Articolo": df_preview["Item Code"].astype(str).str.strip(),
            "Colore": df_preview["Color Code"].astype(str).str.strip().str.zfill(3)
        })
        preview_dfs.append(temp_df)

    preview_final = pd.concat(preview_dfs, ignore_index=True)
    unique_combinations = preview_final[["Articolo", "Colore"]].drop_duplicates()

    st.write("Anteprima Articolo-Colore:")

    selections = {}
    gender_options = ["Seleziona...", "UOMO", "DONNA", "UNISEX"]

    for index, row in unique_combinations.iterrows():
        key = f"{row['Articolo']}-{row['Colore']}"
        remembered = memory_dict.get(key, {})

        preselected_gender = remembered.get("gender", "Seleziona...")
        preselected_base_color = remembered.get("base_color", "")
        preselected_price = remembered.get("price", "")

        col1, col2, col3 = st.columns([1, 1, 1])

        with col1:
            gender_value = st.selectbox(
                f"Gender {key}",
                options=gender_options,
                key=f"gender_{index}",
                index=gender_options.index(preselected_gender) if preselected_gender in gender_options else 0
            )

        with col2:
            base_color_value = st.text_input(
                f"Base Color {key}",
                value=preselected_base_color,
                key=f"base_color_{index}"
            )

        with col3:
            price_value = st.text_input(
                f"Price {key}",
                value=preselected_price,
                key=f"price_{index}"
            )

        selections[(row["Articolo"], row["Colore"])] = {
            "gender": gender_value,
            "base_color": base_color_value.strip(),
            "price": price_value.strip()
        }

    if st.button("Elabora File"):
        if any(v["gender"] == "Seleziona..." for v in selections.values()):
            st.error("Devi selezionare UOMO, DONNA o UNISEX per tutte le combinazioni!")
        else:
            gsheet_data = []
            for _, row in unique_combinations.iterrows():
                sel = selections[(row["Articolo"], row["Colore"])]
                gsheet_data.append((
                    row["Articolo"],
                    row["Colore"],
                    sel["gender"],
                    sel["base_color"],
                    sel["price"]
                ))

            write_to_gsheet(gsheet_data, google_sheet_url)

            updated_memory_dict = get_existing_data(google_sheet_url)

            processed_dfs = []
            for uploaded_file in uploaded_files:
                processed_dfs.append(process_file(uploaded_file, updated_memory_dict))

            final_df = pd.concat(processed_dfs, ignore_index=True)

            uomo_df = final_df[
                final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])]["gender"] == 'UOMO', axis=1)
            ]
            donna_df = final_df[
                final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])]["gender"] == 'DONNA', axis=1)
            ]
            unisex_df = final_df[
                final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])]["gender"] == 'UNISEX', axis=1)
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
