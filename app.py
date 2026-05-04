import streamlit as st
import pandas as pd
from io import BytesIO

def process_excel(file):
    # Membaca file excel
    df = pd.read_excel(file)
    
    # Identifikasi kolom nama panelis (berdasarkan struktur file Anda)
    name_col = 'Nama Panelis (Jika memiliki gelar, mohon disertakan)'
    
    # Mencari kolom berdasarkan kata kunci di dalam kurung siku []
    # Menggunakan regex untuk menangani typo seperti [Keesuaian] atau [Kesesuain]
    kejelasan_cols = [col for col in df.columns if '[Kejelasan]' in col or '[Kejelesan]' in col]
    relevansi_cols = [col for col in df.columns if '[Relevansi]' in col]
    kesesuaian_cols = [col for col in df.columns if '[Kesesuaian]' in col or '[Keesuaian]' in col or '[Kesesuain]' in col]

    # Fungsi untuk membuat dataframe bersih
    def clean_df(cols):
        return df[[name_col] + cols].copy()

    df_kejelasan = clean_df(kejelasan_cols)
    df_relevansi = clean_df(relevansi_cols)
    df_kesesuaian = clean_df(kesesuaian_cols)

    # Simpan ke dalam buffer memory sebagai Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_kejelasan.to_excel(writer, sheet_name='Kejelasan', index=False)
        df_relevansi.to_excel(writer, sheet_name='Relevansi', index=False)
        df_kesesuaian.to_excel(writer, sheet_name='Kesesuaian', index=False)
    
    return output.getvalue()

# UI Streamlit
st.set_page_config(page_title="Expert Judgement Cleaner", layout="centered")
st.title("🛠️ Expert Judgement Data Cleaner")
st.write("Upload file hasil Google Form Anda untuk memisahkan skor berdasarkan kategori.")

uploaded_file = st.file_uploader("Pilih file Excel", type=["xlsx"])

if uploaded_file:
    with st.spinner("Sedang memproses data..."):
        try:
            processed_data = process_excel(uploaded_file)
            
            st.success("Berhasil! Data telah dibersihkan dari kolom 'Keterangan' dan dikelompokkan.")
            
            st.download_button(
                label="📥 Download Excel Hasil Pembersihan",
                data=processed_data,
                file_name="Data_Validasi_Clean.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
