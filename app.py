import streamlit as st
import pandas as pd
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
    rk_file = st.file_uploader("Upload file Rekening Koran", type=["xlsx","xls","csv"])
    db_file = st.file_uploader("Upload file Database", type=["xlsx","xls","csv"])


# =============================
# DETECT HEADER (tanpa asumsi nama kolom)
# =============================
def detect_header(file):

    preview = pd.read_excel(file, header=None, nrows=30)

    for i in range(30):

        row = preview.iloc[i].astype(str).str.lower()

        if any("uraian" in cell for cell in row) \
        or any("description" in cell for cell in row) \
        or any("keterangan" in cell for cell in row) \
        or any("deskripsi" in cell for cell in row):

            return i

    return 0


# =============================
# DETECT KOLOM TEKS TERPANJANG (URAiAN)
# =============================
def detect_text_column(df):

    text_cols = []

    for col in df.columns:

        if df[col].dtype == object:

            avg_len = df[col].astype(str).str.len().mean()

            text_cols.append((col, avg_len))

    if not text_cols:
        return None

    text_cols.sort(key=lambda x: x[1], reverse=True)

    return text_cols[0][0]


# =============================
# DETECT KOLOM DATABASE
# =============================
def detect_db_columns(db):

    db.columns = db.columns.str.strip()

    kode_col = None
    id_col = None

    for col in db.columns:

        name = col.lower()

        if "kode" in name:
            kode_col = col

        if name == "id":
            id_col = col

    return kode_col, id_col


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
        # LOAD FILE
        # =============================
        if mode == "1 File (RK + Database dalam satu file)":

            excel = pd.ExcelFile(rk_file)

            sheet_rk = excel.sheet_names[0]
            sheet_db = excel.sheet_names[1]

            preview = pd.read_excel(excel, sheet_name=sheet_rk, header=None)

            header_row = detect_header(rk_file)

            rk = pd.read_excel(excel, sheet_name=sheet_rk, header=header_row)
            db = pd.read_excel(excel, sheet_name=sheet_db)

        else:

            if rk_file.name.endswith(".csv"):
                rk = pd.read_csv(rk_file)
            else:
                header_row = detect_header(rk_file)
                rk = pd.read_excel(rk_file, header=header_row)

            if db_file is None:
                st.stop()

            if db_file.name.endswith(".csv"):
                db = pd.read_csv(db_file)
            else:
                db = pd.read_excel(db_file)


        # =============================
        # DETECT TEXT COLUMN (URAiAN)
        # =============================
        desc_col = detect_text_column(rk)

        if desc_col is None:

            st.error("Tidak dapat menemukan kolom transaksi")
            st.write("Kolom tersedia:", rk.columns)
            st.stop()


        # =============================
        # DETECT DATABASE COLUMN
        # =============================
        kode_col, id_col = detect_db_columns(db)

        if kode_col is None or id_col is None:

            st.error("Kolom kode unik atau ID tidak ditemukan di database")
            st.write("Kolom database:", db.columns)
            st.stop()


        kode_list = db[kode_col].astype(str).tolist()
        id_list = db[id_col].tolist()


        # =============================
        # GENERATE ID
        # =============================
        rk["ID"] = rk[desc_col].apply(
            lambda x: cari_id(x, kode_list, id_list)
        )


        st.success("ID berhasil digenerate")

        st.dataframe(rk)


        # =============================
        # DOWNLOAD
        # =============================
        output = BytesIO()

        rk.to_excel(output, index=False)

        st.download_button(
            "Download Rekening Koran + ID",
            output.getvalue(),
            "RK_HASIL_ID.xlsx"
        )

    except Exception as e:

        st.error("Terjadi error saat memproses file")
        st.write(e)
