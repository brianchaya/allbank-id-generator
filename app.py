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
        type=["xlsx","csv"]
    )

    db_file = st.file_uploader(
        "Upload file Database",
        type=["xlsx","csv"]
    )


# =====================================
# FUNGSI CARI ID
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
# PROSES
# =====================================
if rk_file:

    try:

        if mode == "1 File (RK + Database dalam satu file)":

            excel = pd.ExcelFile(rk_file)

            rk = pd.read_excel(excel, sheet_name=excel.sheet_names[0])
            db = pd.read_excel(excel, sheet_name=excel.sheet_names[1])

        else:

            if rk_file.name.endswith(".csv"):
                rk = pd.read_csv(rk_file)
            else:
                rk = pd.read_excel(rk_file)

            if db_file.name.endswith(".csv"):
                db = pd.read_csv(db_file)
            else:
                db = pd.read_excel(db_file)


        # =====================================
        # DETECT KOLOM DATABASE
        # =====================================
        db.columns = db.columns.str.strip()

        kode_col = None
        id_col = None

        for col in db.columns:

            if "kode" in col.lower():
                kode_col = col

            if col.lower() == "id":
                id_col = col


        if kode_col is None or id_col is None:

            st.error("Kolom kode unik atau ID tidak ditemukan di database")
            st.stop()


        kode_list = db[kode_col].astype(str).tolist()
        id_list = db[id_col].tolist()


        # =====================================
        # DETECT KOLOM DESKRIPSI TRANSAKSI
        # =====================================
        rk.columns = rk.columns.str.strip()

        desc_col = None

        for col in rk.columns:

            if "uraian" in col.lower():
                desc_col = col

            if "description" in col.lower():
                desc_col = col


        if desc_col is None:

            st.error("Kolom uraian transaksi / description tidak ditemukan")
            st.write("Kolom yang tersedia:", rk.columns)
            st.stop()


        # =====================================
        # GENERATE ID
        # =====================================
        rk["ID"] = rk[desc_col].apply(
            lambda x: cari_id(x, kode_list, id_list)
        )


        st.success("ID berhasil digenerate")

        st.dataframe(rk)


        # =====================================
        # DOWNLOAD
        # =====================================
        output = BytesIO()

        rk.to_excel(output, index=False)

        st.download_button(
            "Download RK dengan ID",
            output.getvalue(),
            "RK_HASIL_ID.xlsx"
        )

    except Exception as e:

        st.error("Terjadi error saat memproses file")
        st.write(e)