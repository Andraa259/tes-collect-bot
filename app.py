import streamlit as st
import pandas as pd
from io import BytesIO

def process_expert_judgement(file):
    # Membaca data
    df = pd.read_excel(file)
    
    # Identifikasi kolom nama secara spesifik sesuai file Anda
    name_col = 'Nama Panelis (Jika memiliki gelar, mohon disertakan)'
    
    # 1. Hapus baris dengan nama 'zaki' (case-insensitive)
    # Ini akan membuat data dimulai dari Jennifer (baris 3 di Excel asli)
    df = df[df[name_col].str.contains('zaki', case=False, na=False) == False]
    
    # 2. Filter kolom berdasarkan kategori [Kejelasan], [Relevansi], [Kesesuaian]
    # Mengabaikan kolom 'Keterangan', 'Email', 'Timestamp', dan 'Universitas'
    kejelasan_cols = [col for col in df.columns if '[Kejelasan]' in col]
    relevansi_cols = [col for col in df.columns if '[Relevansi]' in col]
    kesesuaian_cols = [col for col in df.columns if '[Kesesuaian]' in col]

    # Fungsi pembantu untuk membuat dataframe per sheet
    def create_clean_df(cols):
        return df[[name_col] + cols].copy()

    # 3. Proses Export ke Excel dengan 3 Sheets
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Sheet Kejelasan
        df_kej = create_clean_df(kejelasan_cols)
        df_kej.to_excel(writer, sheet_name='Kejelasan', index=False)
        
        # Sheet Relevansi
        df_rel = create_clean_df(relevansi_cols)
        df_rel.to_excel(writer, sheet_name='Relevansi', index=False)
        
        # Sheet Kesesuaian
        df_kes = create_clean_df(kesesuaian_cols)
        df_kes.to_excel(writer, sheet_name='Kesesuaian', index=False)
        
    return output.getvalue()

# --- UI Streamlit ---
st.set_page_config(page_title="Expert Judgement Cleaner", layout="centered")
st.title("🧼 Expert Judgement Data Cleaner")
st.info("File akan dibersihkan: Nama 'Zaki' dihapus, kolom 'Keterangan' dibuang, dan dikelompokkan per sheet.")

uploaded_file = st.file_uploader("Upload file 'Formulir tanpa judul (Jawaban) (1).xlsx'", type=["xlsx"])

if uploaded_file:
    with st.spinner("Sedang memproses data..."):
        try:
            # Jalankan fungsi pembersihan
            clean_data = process_expert_judgement(uploaded_file)
            
            st.success("✅ Berhasil! Baris 'Zaki' telah dihapus dan data sudah dikelompokkan.")
            
            # Tombol Download
            st.download_button(
                label="📥 Download Data Bersih (Excel)",
                data=clean_data,
                file_name="Data_Validasi_Clean_3.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
