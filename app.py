import streamlit as st
import pandas as pd
import io
import gspread
from google.oauth2.service_account import Credentials

# Opzioni fisse
GENDER_OPTIONS = ["Seleziona...", "UOMO", "DONNA", "UNISEX"]
BASE_COLOR_OPTIONS = [
    "Seleziona...",
    "Black",
    "Blue",
    "Brown",
    "Green",
    "Grey",
    "Metallic",
    "Multicolour",
    "Nude & Neutrals",
    "Pink & Purple",
    "Red",
    "White",
    "Yellow & Orange"
]

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1p84nF9Tq-1ZJgQSEJcgrePLvQyGQ3cjt_1IZP5qPs00/edit?usp=sharing"


def format_taglia(size_us):
    if pd.isna(size_us):
        return ""
    size_str = str(size_us).strip()
    if size_str.endswith(".0"):
        size_str = size_str[:-2]
    return size_str.replace(".5", "+")


def clean_price_value(value):
    if pd.isna(value) or str(value).strip() == "":
        return 0.0
    s = str(value).strip().replace("€", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def clean_multiplier_value(value):
    if pd.isna(value) or str(value).strip() == "":
        return None
    s = str(value).strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def map_country_of_origin(value):
    if pd.isna(value):
        return ""
    v = str(value).strip().upper()
    mapping = {
        "VN": "Vietnam",
        "ID": "Indonesia",
        "KH": "Cambodia"
    }
    return mapping.get(v, "")


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


def get_or_create_gender_worksheet(sheet_url):
    client = connect_to_gsheet()
    sheet = client.open_by_url(sheet_url)

    try:
        worksheet = sheet.worksheet("Gender")
    except gspread.WorksheetNotFound:
        worksheet = sheet.add_worksheet(title="Gender", rows=1000, cols=5)
        worksheet.append_row(["Articolo", "Colore", "Gender", "Base Color", "Price"])

    existing_values = worksheet.get_all_values()
    expected_header = ["Articolo", "Colore", "Gender", "Base Color", "Price"]

    if not existing_values:
        worksheet.append_row(expected_header)
    else:
        header = existing_values[0]
        if header[:5] != expected_header:
            worksheet.update("A1:E1", [expected_header])

    return worksheet


def get_existing_data(sheet_url):
    worksheet = get_or_create_gender_worksheet(sheet_url)
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


def write_to_gsheet(data, sheet_url):
    worksheet = get_or_create_gender_worksheet(sheet_url)
    existing_data = worksheet.get_all_values()

    existing_entries = {
        f"{row[0]}-{row[1]}": idx + 2
        for idx, row in enumerate(existing_data[1:])
        if len(row) >= 2
    }

    batch_updates = []
    rows_to_append = []

    for articolo, colore, gender, base_color, price in data:
        key = f"{articolo}-{colore}"

        if key in existing_entries:
            row_to_update = existing_entries[key]
            batch_updates.append({
                "range": f"C{row_to_update}:E{row_to_update}",
                "values": [[gender, base_color, price]]
            })
        else:
            rows_to_append.append([articolo, colore, gender, base_color, price])

    if batch_updates:
        worksheet.batch_update(batch_updates)

    if rows_to_append:
        worksheet.append_rows(rows_to_append)

    st.success("Dati aggiornati o aggiunti su Google Sheet.")


def read_delivery_items(file):
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
        "Delivery qty.",
        "Country of Origin"
    ]

    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Nel file {file.name} mancano queste colonne: {missing_columns}")
        st.stop()

    return df


def build_unique_combinations(uploaded_files):
    preview_dfs = []

    for uploaded_file in uploaded_files:
        df_preview = read_delivery_items(uploaded_file)

        temp_df = pd.DataFrame({
            "Articolo": df_preview["Item Code"].astype(str).str.strip(),
            "Colore": df_preview["Color Code"].astype(str).str.strip().str.zfill(3)
        })

        preview_dfs.append(temp_df)

    preview_final = pd.concat(preview_dfs, ignore_index=True)
    unique_combinations = preview_final[["Articolo", "Colore"]].drop_duplicates().reset_index(drop=True)
    return unique_combinations


def aggregate_rows_by_barcode(df):
    df = df.copy()

    df["Barcode"] = df["Barcode"].fillna("").astype(str).str.strip()
    df["Qta"] = pd.to_numeric(df["Qta"], errors="coerce").fillna(0)
    df["Costo"] = pd.to_numeric(df["Costo"], errors="coerce").fillna(0)
    df["Retail"] = pd.to_numeric(df["Retail"], errors="coerce").fillna(0)

    df = df[df["Qta"] > 0]
    df = df[df["Barcode"] != ""]

    aggregated = (
        df.groupby("Barcode", as_index=False, dropna=False)
        .agg({
            "Articolo": "first",
            "Descrizione": "first",
            "Categoria": "first",
            "Subcategoria": "first",
            "Colore": "first",
            "Base Color": "first",
            "Made in": "first",
            "Sigla Bimbo": "first",
            "Costo": "first",
            "Retail": "first",
            "Taglia": "first",
            "EAN": "first",
            "Qta": "sum",
            "Materiale": "first",
            "Spec. Materiale": "first",
            "Misure": "first",
            "Scala Taglie": "first",
            "Tacco": "first",
            "Suola": "first",
            "Carryover": "first",
            "HS Code": "first"
        })
    )

    aggregated["Qta"] = aggregated["Qta"].astype(int)
    aggregated["Tot Costo"] = aggregated["Costo"] * aggregated["Qta"]

    aggregated = aggregated[
        [
            "Articolo",
            "Descrizione",
            "Categoria",
            "Subcategoria",
            "Colore",
            "Base Color",
            "Made in",
            "Sigla Bimbo",
            "Costo",
            "Retail",
            "Taglia",
            "Barcode",
            "EAN",
            "Qta",
            "Tot Costo",
            "Materiale",
            "Spec. Materiale",
            "Misure",
            "Scala Taglie",
            "Tacco",
            "Suola",
            "Carryover",
            "HS Code"
        ]
    ]

    return aggregated


def process_file(file, memory_dict, ricarico_value):
    df = read_delivery_items(file)

    keys = (
        df["Item Code"].astype(str).str.strip() + "-" +
        df["Color Code"].astype(str).str.strip().str.zfill(3)
    )

    base_colors = []
    prices = []
    retails = []
    made_in_values = []

    for idx, key in enumerate(keys):
        row_data = memory_dict.get(key, {})

        remembered_base_color = row_data.get("base_color", "")
        if remembered_base_color not in BASE_COLOR_OPTIONS:
            remembered_base_color = ""

        remembered_price = row_data.get("price", "")
        parsed_price = clean_price_value(remembered_price)
        if parsed_price is None:
            parsed_price = 0.0

        retail_value = parsed_price * ricarico_value
        made_in_value = map_country_of_origin(df.iloc[idx]["Country of Origin"])

        base_colors.append("" if remembered_base_color == "Seleziona..." else remembered_base_color)
        prices.append(parsed_price)
        retails.append(retail_value)
        made_in_values.append(made_in_value)

    output_df = pd.DataFrame({
        "Articolo": df["Item Code"].astype(str).str.strip(),
        "Descrizione": df["Item Description"].fillna("").astype(str).str.strip(),
        "Categoria": "CALZATURE",
        "Subcategoria": "Sneakers",
        "Colore": df["Color Code"].astype(str).str.strip().str.zfill(3),
        "Base Color": base_colors,
        "Made in": made_in_values,
        "Sigla Bimbo": "",
        "Costo": prices,
        "Retail": retails,
        "Taglia": df["US Size"].apply(format_taglia),
        "Barcode": df["EAN Code"].fillna("").astype(str).str.strip(),
        "EAN": "",
        "Qta": pd.to_numeric(df["Delivery qty."], errors="coerce").fillna(0).astype(int),
        "Tot Costo": 0,
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
    aggregated_df = aggregate_rows_by_barcode(output_df)
    return aggregated_df


def write_data_in_chunks(writer, df, stagione, data_inizio, data_fine, ricarico):
    num_chunks = len(df) // 50 + (1 if len(df) % 50 > 0 else 0)

    for i in range(num_chunks):
        chunk_df = df[i * 50:(i + 1) * 50]
        sheet_name = f"Foglio{i+1}"
        start_row = 9

        chunk_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)

        worksheet = writer.sheets[sheet_name]

        worksheet.write("A1", "STAGIONE:")
        worksheet.write("B1", stagione)
        worksheet.write("A2", "TIPO:")
        worksheet.write("B2", "ACCESSORI")
        worksheet.write("A3", "DATA INIZIO:")
        worksheet.write("B3", data_inizio.strftime("%d/%m/%Y"))
        worksheet.write("A4", "DATA FINE:")
        worksheet.write("B4", data_fine.strftime("%d/%m/%Y"))
        worksheet.write("A5", "RICARICO:")
        worksheet.write("B5", ricarico)

        text_format = writer.book.add_format({"num_format": "@"})
        number_format = writer.book.add_format({"num_format": "#,##0.00"})

        worksheet.set_column("A:A", 18)
        worksheet.set_column("B:B", 35)
        worksheet.set_column("C:C", 15)
        worksheet.set_column("D:D", 18)
        worksheet.set_column("E:E", 12)
        worksheet.set_column("F:F", 18)
        worksheet.set_column("G:G", 15)
        worksheet.set_column("H:H", 12)
        worksheet.set_column("I:I", 12, number_format)
        worksheet.set_column("J:J", 12, number_format)
        worksheet.set_column("K:K", 10)
        worksheet.set_column("L:L", 20, text_format)
        worksheet.set_column("M:M", 12)
        worksheet.set_column("N:N", 10)
        worksheet.set_column("O:O", 12, number_format)
        worksheet.set_column("P:W", 15)

        last_data_row = len(chunk_df) + start_row
        empty_row = last_data_row + 1
        worksheet.write(f"N{empty_row}", "")
        worksheet.write(f"O{empty_row}", "")

        total_row = empty_row + 2
        worksheet.write_formula(
            f"N{total_row}",
            f"=SUM(N{start_row+2}:N{last_data_row + 1})"
        )
        worksheet.write_formula(
            f"O{total_row}",
            f"=SUM(O{start_row+2}:O{last_data_row + 1})",
            number_format
        )


st.title("Pkg Asics Xmag Lineare")

stagione = st.text_input("Inserisci STAGIONE")
data_inizio = st.date_input("Inserisci DATA INIZIO")
data_fine = st.date_input("Inserisci DATA FINE")
ricarico = st.text_input("Inserisci RICARICO", value="2")

uploaded_files = st.file_uploader(
    "Scegli i file Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files and stagione and data_inizio and data_fine:
    memory_dict = get_existing_data(GOOGLE_SHEET_URL)
    unique_combinations = build_unique_combinations(uploaded_files)

    st.write("Anteprima Articolo-Colore:")

    selections = {}

    for index, row in unique_combinations.iterrows():
        key = f"{row['Articolo']}-{row['Colore']}"
        remembered = memory_dict.get(key, {})

        preselected_gender = remembered.get("gender", "Seleziona...")
        if preselected_gender not in GENDER_OPTIONS:
            preselected_gender = "Seleziona..."

        preselected_base_color = remembered.get("base_color", "Seleziona...")
        if preselected_base_color not in BASE_COLOR_OPTIONS:
            preselected_base_color = "Seleziona..."

        preselected_price = remembered.get("price", "")

        col1, col2, col3 = st.columns([1, 1, 1])

        with col1:
            gender_value = st.selectbox(
                f"Gender {key}",
                options=GENDER_OPTIONS,
                key=f"gender_{index}",
                index=GENDER_OPTIONS.index(preselected_gender)
            )

        with col2:
            base_color_value = st.selectbox(
                f"Base Color {key}",
                options=BASE_COLOR_OPTIONS,
                key=f"base_color_{index}",
                index=BASE_COLOR_OPTIONS.index(preselected_base_color)
            )

        with col3:
            price_value = st.text_input(
                f"Price {key}",
                value=preselected_price,
                key=f"price_{index}"
            )

        selections[(row["Articolo"], row["Colore"])] = {
            "gender": gender_value,
            "base_color": base_color_value,
            "price": price_value.strip()
        }

    if st.button("Elabora File"):
        ricarico_value = clean_multiplier_value(ricarico)
        if ricarico_value is None:
            st.error("RICARICO non valido.")
            st.stop()

        if any(v["gender"] == "Seleziona..." for v in selections.values()):
            st.error("Devi selezionare UOMO, DONNA o UNISEX per tutte le combinazioni!")
            st.stop()

        if any(v["base_color"] == "Seleziona..." for v in selections.values()):
            st.error("Devi selezionare un Base Color per tutte le combinazioni!")
            st.stop()

        invalid_prices = []
        for (articolo, colore), values in selections.items():
            cleaned = clean_price_value(values["price"])
            if values["price"] != "" and cleaned is None:
                invalid_prices.append(f"{articolo}-{colore}")

        if invalid_prices:
            st.error("Prezzo non valido per: " + ", ".join(invalid_prices))
            st.stop()

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

        write_to_gsheet(gsheet_data, GOOGLE_SHEET_URL)
        updated_memory_dict = get_existing_data(GOOGLE_SHEET_URL)

        processed_dfs = []
        for uploaded_file in uploaded_files:
            processed_dfs.append(process_file(uploaded_file, updated_memory_dict, ricarico_value))

        final_df = pd.concat(processed_dfs, ignore_index=True)

        uomo_df = final_df[
            final_df.apply(
                lambda x: selections.get((x["Articolo"], x["Colore"]), {}).get("gender") == "UOMO",
                axis=1
            )
        ].copy()

        donna_df = final_df[
            final_df.apply(
                lambda x: selections.get((x["Articolo"], x["Colore"]), {}).get("gender") == "DONNA",
                axis=1
            )
        ].copy()

        unisex_df = final_df[
            final_df.apply(
                lambda x: selections.get((x["Articolo"], x["Colore"]), {}).get("gender") == "UNISEX",
                axis=1
            )
        ].copy()

        uomo_output = io.BytesIO()
        donna_output = io.BytesIO()
        unisex_output = io.BytesIO()

        if not uomo_df.empty:
            with pd.ExcelWriter(uomo_output, engine="xlsxwriter") as writer_uomo:
                write_data_in_chunks(writer_uomo, uomo_df, stagione, data_inizio, data_fine, ricarico)
            st.download_button(
                label="Download File UOMO",
                data=uomo_output.getvalue(),
                file_name="uomo_processed_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        if not donna_df.empty:
            with pd.ExcelWriter(donna_output, engine="xlsxwriter") as writer_donna:
                write_data_in_chunks(writer_donna, donna_df, stagione, data_inizio, data_fine, ricarico)
            st.download_button(
                label="Download File DONNA",
                data=donna_output.getvalue(),
                file_name="donna_processed_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        if not unisex_df.empty:
            with pd.ExcelWriter(unisex_output, engine="xlsxwriter") as writer_unisex:
                write_data_in_chunks(writer_unisex, unisex_df, stagione, data_inizio, data_fine, ricarico)
            st.download_button(
                label="Download File UNISEX",
                data=unisex_output.getvalue(),
                file_name="unisex_processed_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
