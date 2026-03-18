import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.title("All Bank ID Generator")

mode = st.radio(
    "Pilih jenis file",
    ["1 File (RK + Database)", "2 File (RK dan Database terpisah)"]
)

rk_file = None
db_file = None

if mode == "1 File (RK + Database)":
    rk_file = st.file_uploader("Upload file RK + Database", type=["xlsx"])
else:
    rk_file = st.file_uploader("Upload Rekening Koran", type=["xlsx"])
    db_file = st.file_uploader("Upload Database", type=["xlsx"])


# =====================================
# DETECT HEADER
# =====================================
def detect_header(df):

    for i in range(min(20, len(df))):

        row = df.iloc[i].astype(str).str.lower()

        if any("uraian" in c for c in row) \
        or any("description" in c for c in row) \
        or any("keterangan" in c for c in row) \
        or any("deskripsi" in c for c in row):

            return i

    return 0


# =====================================
# DETECT RK SHEET
# =====================================
def detect_rk_sheet(excel):

    for sheet in excel.sheet_names:

        df = pd.read_excel(excel, sheet_name=sheet, nrows=20)

        cols = [str(c).lower() for c in df.columns]

        if any("uraian" in c for c in cols) \
        or any("description" in c for c in cols) \
        or any("keterangan" in c for c in cols):

            return sheet

    return excel.sheet_names[0]


# =====================================
# DETECT DATABASE SHEET
# =====================================
def detect_db_sheet(excel):

    for sheet in excel.sheet_names:

        df = pd.read_excel(excel, sheet_name=sheet, nrows=10)

        cols = [str(c).lower() for c in df.columns]

        if "id" in cols:
            return sheet

    return excel.sheet_names[-1]


# =====================================
# DETECT TRANSACTION COLUMN
# =====================================
def detect_transaction_col(df):

    for col in df.columns:

        name = str(col).lower()

        if "uraian" in name \
        or "description" in name \
        or "keterangan" in name \
        or "deskripsi" in name:

            return col

    # fallback → kolom teks paling panjang
    lengths = df.astype(str).apply(lambda x: x.str.len().mean())

    return lengths.idxmax()


# =====================================
# DETECT DB COLUMN
# =====================================
def detect_db_columns(db):

    kode_col = None
    id_col = None

    for col in db.columns:

        name = str(col).lower()

        if "kode" in name:
            kode_col = col

        if name == "id":
            id_col = col

    return kode_col, id_col


# =====================================
# FAST SEARCH
# =====================================
def generate_ids(text_series, kode_list, id_list):

    # gabung & sort dari yang paling panjang
    pairs = list(zip(kode_list, id_list))
    pairs.sort(key=lambda x: len(str(x[0])), reverse=True)

    results = []

    for text in text_series.astype(str).str.lower():

        found = None

        for kode, id_val in pairs:

            if str(kode).lower() in text:

                found = id_val
                break

        results.append(found)

    return results


# =====================================
# MAIN PROCESS
# =====================================
if mode == "1 File (RK + Database)" and rk_file:

    try:

        excel = pd.ExcelFile(rk_file)

        rk_sheet = detect_rk_sheet(excel)
        db_sheet = detect_db_sheet(excel)

        preview = pd.read_excel(excel, sheet_name=rk_sheet, header=None)

        header_row = detect_header(preview)

        rk = pd.read_excel(excel, sheet_name=rk_sheet, header=header_row)
        db = pd.read_excel(excel, sheet_name=db_sheet)

        wb = load_workbook(rk_file)
        ws = wb[rk_sheet]

    except Exception as e:

        st.error(e)
        st.stop()


elif mode == "2 File (RK dan Database terpisah)" and rk_file and db_file:

    try:

        excel_rk = pd.ExcelFile(rk_file)
        rk_sheet = detect_rk_sheet(excel_rk)

        preview = pd.read_excel(excel_rk, sheet_name=rk_sheet, header=None)

        header_row = detect_header(preview)

        rk = pd.read_excel(excel_rk, sheet_name=rk_sheet, header=header_row)

        excel_db = pd.ExcelFile(db_file)
        db_sheet = detect_db_sheet(excel_db)

        db = pd.read_excel(excel_db, sheet_name=db_sheet)

        wb = load_workbook(rk_file)
        ws = wb[rk_sheet]

    except Exception as e:

        st.error(e)
        st.stop()

else:

    st.stop()


# =====================================
# DETECT COLUMN
# =====================================
desc_col = detect_transaction_col(rk)

kode_col, id_col = detect_db_columns(db)

if kode_col is None or id_col is None:

    st.error("Kolom kode unik / ID tidak ditemukan di database")
    st.stop()


kode_list = db[kode_col].astype(str).tolist()
id_list = db[id_col].tolist()


# =====================================
# GENERATE IDS
# =====================================
ids = generate_ids(rk[desc_col], kode_list, id_list)

if "ID" in rk.columns:
    rk["ID"] = ids
else:
    rk.insert(len(rk.columns), "ID", ids)

st.subheader("Preview Hasil (Full Data)")
st.dataframe(rk)

missing_count = rk["ID"].isna().sum()
st.write(f"Jumlah ID tidak terisi: {missing_count}")

# =====================================
# WRITE BACK TO EXCEL (KEEP FORMAT)
# =====================================
header_row_excel = None
id_col_excel = None

for r in range(1, 11):

    for c in range(1, ws.max_column+1):

        val = ws.cell(r,c).value

        if val and str(val).lower() == "id":

            header_row_excel = r
            id_col_excel = c


if header_row_excel is None:

    header_row_excel = 1
    id_col_excel = ws.max_column + 1

    ws.cell(header_row_excel, id_col_excel).value = "ID"


from openpyxl.styles import PatternFill

red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

for i,val in enumerate(rk["ID"]):

    cell = ws.cell(header_row_excel+i, id_col_excel)
    cell.value = val

    if pd.isna(val):
        cell.fill = red_fill


# =====================================
# DOWNLOAD
# =====================================
output = BytesIO()

wb.save(output)

st.success("ID berhasil digenerate")

st.download_button(
    "Download RK dengan ID",
    output.getvalue(),
    "RK_HASIL_ID.xlsx"
)
