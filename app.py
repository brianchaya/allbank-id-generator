import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.title("Auto Generate ID dari Database")


mode = st.radio(
    "Pilih jenis file",
    ["1 File (RK + Database dalam satu file)", "2 File (RK dan Database terpisah)"]
)

rk_file = None
db_file = None

if mode == "1 File (RK + Database dalam satu file)":
    rk_file = st.file_uploader("Upload file RK + Database", type=["xlsx"])
else:
    rk_file = st.file_uploader("Upload file Rekening Koran", type=["xlsx"])
    db_file = st.file_uploader("Upload file Database", type=["xlsx"])


# =============================
# DETECT SHEET RK
# =============================
def detect_rk_sheet(excel):

    for sheet in excel.sheet_names:

        df = pd.read_excel(excel, sheet_name=sheet, nrows=20)

        cols = [str(c).lower() for c in df.columns]

        if any("uraian" in c for c in cols) or \
           any("description" in c for c in cols) or \
           any("keterangan" in c for c in cols):

            return sheet

    return excel.sheet_names[0]


# =============================
# DETECT SHEET DATABASE
# =============================
def detect_db_sheet(excel):

    for sheet in excel.sheet_names:

        df = pd.read_excel(excel, sheet_name=sheet, nrows=10)

        cols = [str(c).lower() for c in df.columns]

        if "id" in cols:
            return sheet

    return excel.sheet_names[0]


# =============================
# SEARCH ID
# =============================
def cari_id(text, kode_list, id_list):

    if pd.isna(text):
        return None

    text = str(text).lower()

    for kode, id_val in zip(kode_list, id_list):

        if str(kode).lower() in text:
            return id_val

    return None


# =============================
# MAIN PROCESS
# =============================
if rk_file:

    try:

        # =============================
        # LOAD RK & DB
        # =============================
        if mode == "1 File (RK + Database dalam satu file)":

            excel = pd.ExcelFile(rk_file)

            rk_sheet = detect_rk_sheet(excel)
            db_sheet = detect_db_sheet(excel)

            rk = pd.read_excel(excel, sheet_name=rk_sheet)
            db = pd.read_excel(excel, sheet_name=db_sheet)

            wb = load_workbook(rk_file)
            ws = wb[rk_sheet]

        else:

            excel_rk = pd.ExcelFile(rk_file)

            rk_sheet = detect_rk_sheet(excel_rk)

            rk = pd.read_excel(excel_rk, sheet_name=rk_sheet)

            excel_db = pd.ExcelFile(db_file)

            db_sheet = detect_db_sheet(excel_db)

            db = pd.read_excel(excel_db, sheet_name=db_sheet)

            wb = load_workbook(rk_file)
            ws = wb[rk_sheet]


        # =============================
        # DETECT KOLOM DATABASE
        # =============================
        db.columns = db.columns.str.strip()

        kode_col = None
        id_col = None

        for col in db.columns:

            name = col.lower()

            if "kode" in name:
                kode_col = col

            if name == "id":
                id_col = col


        kode_list = db[kode_col].astype(str).tolist()
        id_list = db[id_col].tolist()


        # =============================
        # DETECT KOLOM TRANSAKSI
        # =============================
        desc_col = None

        for col in rk.columns:

            name = col.lower()

            if "uraian" in name or \
               "description" in name or \
               "keterangan" in name or \
               "deskripsi" in name:

                desc_col = col


        # =============================
        # CARI BARIS HEADER
        # =============================
        header_row = None

        for r in range(1, 11):

            for c in range(1, ws.max_column+1):

                val = ws.cell(r,c).value

                if val and str(val).lower() == "id":

                    header_row = r
                    id_col_excel = c

        if header_row is None:

            header_row = 1
            id_col_excel = ws.max_column + 1

            ws.cell(header_row, id_col_excel).value = "ID"


        # =============================
        # GENERATE ID
        # =============================
        desc_index = rk.columns.get_loc(desc_col)

        for i in range(len(rk)):

            row_excel = header_row + 1 + i

            text = rk.iloc[i, desc_index]

            id_val = cari_id(text, kode_list, id_list)

            ws.cell(row_excel, id_col_excel).value = id_val


        st.success("ID berhasil digenerate")


        # =============================
        # SAVE FILE TANPA MERUSAK FORMAT
        # =============================
        output = BytesIO()

        wb.save(output)

        st.download_button(
            "Download Rekening Koran + ID",
            output.getvalue(),
            "RK_HASIL_ID.xlsx"
        )

    except Exception as e:

        st.error("Terjadi error saat memproses file")
        st.write(e)
