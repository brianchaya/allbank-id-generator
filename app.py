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

    rk_file = st.file_uploader(
        "Upload file RK + Database",
        type=["xlsx"]
    )

else:

    rk_file = st.file_uploader(
        "Upload file Rekening Koran",
        type=["xlsx","xls","csv"]
    )

    db_file = st.file_uploader(
        "Upload file Database",
        type=["xlsx","xls","csv"]
    )


# =====================================
# DETECT HEADER
# =====================================
def detect_header(file):

    preview = pd.read_excel(file, header=None, nrows=20)

    for i in range(20):

        row = preview.iloc[i].astype(str).str.lower()

        if any("uraian" in cell for cell in row) or \
           any("description" in cell for cell in row) or \
           any("keterangan" in cell for cell in row):

            return i

    return 0


# =====================================
# DETECT KOLOM URAIAN
# =====================================
def detect_desc_column(df):

    df.columns = df.columns.str.strip()

    for col in df.columns:

        name = col.lower()

        if "uraian" in name:
            return col

        if "description" in name:
            return col

        if "keterangan" in name:
            return col

    return None


# =====================================
# DETECT KOLOM DATABASE
# =====================================
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


# =====================================
# CARI ID
# =====================================
def cari_id(text, kode_list, id_list):

    if pd.isna(text):
        return None

    text = str(text).lower()

    for kode, id_val in zip(kode_list, id_list):

        if str(kode).lower() in text:
            return id_val

    return None


# =====================================
# MAIN PROCESS
# =====================================
if rk_file:

    try:

        # =============================
        # LOAD FILE
        # =============================
        if mode == "1 File (RK + Database dalam satu file)":

            excel = pd.ExcelFile(rk_file)

            rk = pd.read_excel(excel, sheet_name=excel.sheet_names[0])
            db = pd.read_excel(excel, sheet_name=excel.sheet_names[1])

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
        # DETECT COLUMN RK
        # =============================
        desc_col = detect_desc_column(rk)

        if desc_col is None:

            st.error("Kolom uraian transaksi / description tidak ditemukan")
            st.write("Kolom yang tersedia:", rk.columns)
            st.stop()


        # =============================
        # DETECT COLUMN DATABASE
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