import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
from openpyxl.styles import PatternFill

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


# ===============================
# DETECT HEADER
# ===============================
def detect_header(df):
    for i in range(min(20, len(df))):
        row = df.iloc[i].astype(str).str.lower()
        if any(k in c for c in row for k in ["uraian","description","keterangan","deskripsi"]):
            return i
    return 0


# ===============================
# DETECT KOLOM
# ===============================
def detect_transaction_col(df):
    for col in df.columns:
        if any(k in str(col).lower() for k in ["uraian","description","keterangan","deskripsi"]):
            return col
    return df.columns[0]


def detect_db_columns(db):
    kode_col, id_col = None, None
    for col in db.columns:
        name = str(col).strip().lower()
        if "kode" in name:
            kode_col = col
        if name == "id":
            id_col = col
    return kode_col, id_col


# ===============================
# MATCHING
# ===============================
def generate_ids(text_series, kode_list, id_list):
    pairs = list(zip(kode_list, id_list))
    pairs.sort(key=lambda x: len(str(x[0])), reverse=True)

    result = []
    for text in text_series.astype(str).str.lower():
        found = None
        for kode, id_val in pairs:
            if str(kode).lower() in text:
                found = id_val
                break
        result.append(found)
    return result


# ===============================
# LOAD
# ===============================
if mode == "1 File (RK + Database)" and rk_file:
    excel = pd.ExcelFile(rk_file)

    sheet = excel.sheet_names[0]
    preview = pd.read_excel(excel, sheet_name=sheet, header=None)
    header_row = detect_header(preview)

    rk = pd.read_excel(excel, sheet_name=sheet, header=header_row)
    db = pd.read_excel(excel, sheet_name=excel.sheet_names[-1])

elif mode == "2 File (RK dan Database terpisah)" and rk_file and db_file:
    excel_rk = pd.ExcelFile(rk_file)

    sheet = excel_rk.sheet_names[0]
    preview = pd.read_excel(excel_rk, sheet_name=sheet, header=None)
    header_row = detect_header(preview)

    rk = pd.read_excel(excel_rk, sheet_name=sheet, header=header_row)
    db = pd.read_excel(db_file)

else:
    st.stop()


# ===============================
# PROCESS
# ===============================
desc_col = detect_transaction_col(rk)
kode_col, id_col = detect_db_columns(db)

if kode_col is None or id_col is None:
    st.error("Kolom kode / ID tidak ditemukan")
    st.stop()

ids = generate_ids(
    rk[desc_col],
    db[kode_col].astype(str).tolist(),
    db[id_col].tolist()
)

# pastikan kolom ID konsisten
if "ID" in [str(c).strip().upper() for c in rk.columns]:
    for col in rk.columns:
        if str(col).strip().upper() == "ID":
            rk[col] = ids
            col_used = col
            break
else:
    rk["ID"] = ids
    col_used = "ID"


# ===============================
# PREVIEW
# ===============================
st.subheader("Preview")
st.dataframe(rk)

missing = rk[col_used].isna().sum()
st.write(f"Jumlah ID tidak terisi: {missing}")


# ===============================
# TULIS ULANG TOTAL (ANTI ERROR)
# ===============================
wb = load_workbook(rk_file)
ws = wb.active

red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

# posisi header excel
header_excel = header_row + 1

# cari index kolom ID di pandas
id_index = list(rk.columns).index(col_used) + 1  # excel 1-based

# set header
ws.cell(header_excel, id_index).value = "ID"

# isi data
for i, val in enumerate(rk[col_used].values):
    row_excel = header_excel + 1 + i

    cell = ws.cell(row_excel, id_index)
    cell.value = val

    if pd.isna(val):
        cell.fill = red_fill


# ===============================
# DOWNLOAD
# ===============================
output = BytesIO()
wb.save(output)

st.success("Selesai")

st.download_button(
    "Download",
    output.getvalue(),
    "RK_HASIL.xlsx"
)
