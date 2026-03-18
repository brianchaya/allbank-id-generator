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
        if any(k in c for c in row for k in ["uraian","description","keterangan","deskripsi"]):
            return i
    return 0


# =====================================
# DETECT RK SHEET
# =====================================
def detect_rk_sheet(excel):
    for sheet in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet, nrows=20)
        cols = [str(c).lower() for c in df.columns]
        if any(k in c for c in cols for k in ["uraian","description","keterangan"]):
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
        if any(k in name for k in ["uraian","description","keterangan","deskripsi"]):
            return col

    lengths = df.astype(str).apply(lambda x: x.str.len().mean())
    return lengths.idxmax()


# =====================================
# DETECT DB COLUMN
# =====================================
def detect_db_columns(db):
    kode_col = None
    id_col = None

    for col in db.columns:
        name = str(col).strip().lower()
        if "kode" in name:
            kode_col = col
        if name == "id":
            id_col = col

    return kode_col, id_col


# =====================================
# MATCHING
# =====================================
def generate_ids(text_series, kode_list, id_list):
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
# LOAD DATA
# =====================================
if mode == "1 File (RK + Database)" and rk_file:

    excel = pd.ExcelFile(rk_file)
    rk_sheet = detect_rk_sheet(excel)
    db_sheet = detect_db_sheet(excel)

    preview = pd.read_excel(excel, sheet_name=rk_sheet, header=None)
    header_row = detect_header(preview)

    rk = pd.read_excel(excel, sheet_name=rk_sheet, header=header_row)
    db = pd.read_excel(excel, sheet_name=db_sheet)

    wb = load_workbook(rk_file)
    ws = wb[rk_sheet]


elif mode == "2 File (RK dan Database terpisah)" and rk_file and db_file:

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

else:
    st.stop()


# =====================================
# PREP
# =====================================
desc_col = detect_transaction_col(rk)
kode_col, id_col = detect_db_columns(db)

if kode_col is None or id_col is None:
    st.error("Kolom kode / ID tidak ditemukan")
    st.stop()

kode_list = db[kode_col].astype(str).tolist()
id_list = db[id_col].tolist()


# =====================================
# GENERATE
# =====================================
ids = generate_ids(rk[desc_col], kode_list, id_list)

# DETECT KOLOM ID ASLI
id_col_name = None
for col in rk.columns:
    if str(col).strip().lower() == "id":
        id_col_name = col
        break

# ISI KE DATAFRAME (INI YG NGARUH KE PREVIEW)
if id_col_name:
    rk[id_col_name] = ids
    col_used = id_col_name
else:
    rk["ID"] = ids
    col_used = "ID"


# =====================================
# PREVIEW
# =====================================
st.subheader("Preview Hasil")
st.dataframe(rk)

missing = rk[col_used].isna().sum()
st.write(f"Jumlah ID tidak terisi: {missing}")


# =====================================
# WRITE KE EXCEL (FIX TOTAL)
# =====================================

# posisi header excel = header pandas + 1
header_excel = header_row + 1

# cari kolom ID di excel (baris header)
id_col_excel = None

for c in range(1, ws.max_column + 1):
    val = ws.cell(header_excel, c).value
    if val and str(val).strip().lower() == "id":
        id_col_excel = c
        break

# kalau ga ada → tambah di kanan
if id_col_excel is None:
    id_col_excel = ws.max_column + 1
    ws.cell(header_excel, id_col_excel).value = "ID"
else:
    ws.cell(header_excel, id_col_excel).value = "ID"


from openpyxl.styles import PatternFill
red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

# tulis data (NO OFFSET ERROR LAGI)
for i, val in enumerate(rk[col_used].values):
    excel_row = header_excel + 1 + i  # ini fix utama

    cell = ws.cell(excel_row, id_col_excel)
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
