import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Lendable Limit Automation", layout="wide")

# ===============================================================
#               📊 LENDABLE LIMIT AUTOMATION (Streamlit)
# ===============================================================

st.title("📊 Lendable Limit Automation (PEI - RMCC)")
st.caption("Upload file Excel perhitungan dan sistem akan otomatis memproses serta menampilkan hasilnya.")

# ===============================================================
# Fungsi bantu
# ===============================================================

def detect_header_row(df_raw):
    """Coba cari baris mana yang paling cocok jadi header"""
    for i in range(3):
        sample = pd.read_excel(df_raw, header=i, nrows=1)
        if 'NAMA EFEK' in sample.columns or 'KODE EFEK' in sample.columns:
            return i
    return 0  # fallback default ke baris pertama


def process_and_style_conc_limit(df):
    """Proses data concentration limit dan cek kolom"""

    try:
        # --- Normalisasi nama kolom biar aman ---
        df.columns = df.columns.str.strip().str.upper()

        # --- Definisi kolom yang diperlukan ---
        FINAL_COLS = [
            'KODE EFEK', 'NAMA EFEK', 'HAIRCUT KPEI LAMA', 'HAIRCUT KPEI BARU',
            'HAIRCUT PEI USULAN DIVISI', 'CLOSING PRICE', 'LISTED SHARES',
            'FREE FLOAT (DALAM LEMBAR)', 'PERBANDINGAN DENGAN LISTED SHARES (SESUAI PERHITUNGAN)',
            'PERBANDINGAN DENGAN FREE FLOAT (SESUAI PERHITUNGAN)',
            'CONCENTRATION LIMIT (Sesuai Perhitungan)',
            'CONCENTRATION LIMIT KARENA SAHAM MARJIN BARU',
            'CONCENTRATION LIMIT TERKENA % LISTED SHARES',
            'CONCENTRATION LIMIT TERKENA % FREEFLOAT',
            'CONCENTRATION LIMIT FINAL RMCC',
            'SAHAM MARJIN BARU?', 'UMA', 'KETERANGAN', 'KETERANGAN UMA'
        ]

        # --- Cek apakah semua kolom tersedia ---
        missing_cols = [col for col in FINAL_COLS if col not in df.columns]
        if missing_cols:
            st.error(f"❌ Kolom berikut tidak ditemukan di file Excel:\n{', '.join(missing_cols)}")
            st.write("Kolom yang ditemukan di file kamu adalah:")
            st.write(list(df.columns))
            st.stop()

        # --- Proses data ---
        df_result = df[FINAL_COLS].copy()

        # Styling sederhana (contoh)
        df_result['HAIRCUT KPEI LAMA'] = df_result['HAIRCUT KPEI LAMA'].fillna(0)
        df_result['HAIRCUT KPEI BARU'] = df_result['HAIRCUT KPEI BARU'].fillna(0)
        df_result['HAIRCUT PEI USULAN DIVISI'] = df_result['HAIRCUT PEI USULAN DIVISI'].fillna(0)

        return df_result

    except Exception as e:
        st.error(f"❌ CL: Gagal memproses Concentration Limit: {e}")
        st.stop()

# ===============================================================
# Upload Section
# ===============================================================

uploaded_file = st.file_uploader("📁 Upload file Excel (xlsx/xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Deteksi header otomatis
        header_row = detect_header_row(uploaded_file)
        df = pd.read_excel(uploaded_file, header=header_row, engine='openpyxl')

        st.success(f"✅ File berhasil dibaca! (Header terdeteksi di baris ke-{header_row + 1})")
        st.write("📄 **Preview data:**")
        st.dataframe(df.head())

        # Proses data
        df_result = process_and_style_conc_limit(df)

        st.divider()
        st.subheader("📊 Hasil Concentration Limit")
        st.dataframe(df_result)

        # Tombol download hasil
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='Result')
        st.download_button(
            label="⬇️ Download Hasil Excel",
            data=buffer.getvalue(),
            file_name=f"Lendable_Limit_Result_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Gagal membaca file: {e}")

else:
    st.info("⬆️ Silakan upload file Excel untuk mulai memproses.")