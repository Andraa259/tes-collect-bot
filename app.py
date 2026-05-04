import streamlit as st
import pandas as pd
from io import BytesIO

def process_data_psp(file):
    # Membaca file
    df = pd.read_excel(file)
    
    # Identifikasi kolom Nama (berdasarkan struktur file Aitem PSP.xlsx)
    # Kolom ini: 'NAMA (Jika memiliki gelar, tolong di sertakan)'
    name_col = [col for col in df.columns if 'NAMA' in col.upper()][0]
    
    # Mencari kolom skor berdasarkan kategori di dalam kurung siku
    kejelasan_cols = [col for col in df.columns if '[Kejelasan]' in col]
    relevansi_cols = [col for col in df.columns if '[Relevansi]' in col]
    kesesuaian_cols = [col for col in df.columns if '[Kesesuaian]' in col]

    # Fungsi untuk membersihkan dan mengambil Nama + Skor saja
    def clean_sheet(cols):
        return df[[name_col] + cols].copy()

    # Ekspor ke Excel dengan multiple sheets
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        clean_sheet(kejelasan_cols).to_excel(writer, sheet_name='Kejelasan', index=False)
        clean_sheet(relevansi_cols).to_excel(writer, sheet_name='Relevansi', index=False)
        clean_sheet(kesesuaian_cols).to_excel(writer, sheet_name='Kesesuaian', index=False)
    
    return output.getvalue()

# Antarmuka Streamlit
st.set_page_config(page_title="PSP Data Cleaner", page_icon="📑")
st.title("📑 PSP Item Response Cleaner")
st.write("Program ini akan membersihkan file **Aitem PSP.xlsx** dan memisahkan Skor berdasarkan kategori.")

uploaded_file = st.file_uploader("Unggah file Aitem PSP.xlsx", type=["xlsx"])

if uploaded_file:
    with st.spinner("Sedang memproses..."):
        try:
            processed_file = process_data_psp(uploaded_file)
            
            st.success("Pembersihan selesai! Email, Timestamp, dan Universitas telah dihapus.")
            
            st.download_button(
                label="📥 Download File Terpilih (Nama & Skor Saja)",
                data=processed_file,
                file_name="Aitem_PSP_Cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan teknis: {e}")
