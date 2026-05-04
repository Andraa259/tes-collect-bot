import streamlit as st
import pandas as pd
from io import BytesIO

def process_expert_judgement_v2(file):
    # Membaca data
    df = pd.read_excel(file)
    
    # Identifikasi kolom nama
    name_col = 'Nama Panelis (Jika memiliki gelar, mohon disertakan)'
    
    # 1. Hapus baris 'zaki' (case-insensitive)
    # Filter ini memastikan baris zaki hilang dan data mulai dari baris berikutnya (Jennifer)
    df = df[df[name_col].str.contains('zaki', case=False, na=False) == False]
    
    # 2. Filter kolom dengan logika "Partial Match" untuk menangani spasi (misal: [Kejelasan ])
    # Kode ini akan mengambil semua aitem (total 54) tanpa ada yang tertinggal
    kejelasan_cols = [col for col in df.columns if '[Kejelasan' in col]
    relevansi_cols = [col for col in df.columns if '[Relevansi' in col]
    kesesuaian_cols = [col for col in df.columns if '[Kesesuaian' in col]

    def create_clean_df(cols):
        return df[[name_col] + cols].copy()

    # 3. Proses Export
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        create_clean_df(kejelasan_cols).to_excel(writer, sheet_name='Kejelasan', index=False)
        create_clean_df(relevansi_cols).to_excel(writer, sheet_name='Relevansi', index=False)
        create_clean_df(kesesuaian_cols).to_excel(writer, sheet_name='Kesesuaian', index=False)
        
    return output.getvalue(), len(kejelasan_cols)

# --- UI Streamlit ---
st.title("🚀 Expert Judgement Cleaner (Fix 54 Items)")
st.markdown("Pembersihan otomatis untuk file: **Formulir tanpa judul (Jawaban) (1).xlsx**")

uploaded_file = st.file_uploader("Upload File", type=["xlsx"])

if uploaded_file:
    clean_data, total_found = process_expert_judgement_v2(uploaded_file)
    
    # Menampilkan konfirmasi jumlah aitem agar kamu yakin
    st.success(f"Berhasil memproses **{total_found} aitem**! Baris 'Zaki' telah dibuang.")
    
    st.download_button(
        label="📥 Download Data Bersih (54 Aitem)",
        data=clean_data,
        file_name="Data_Validasi_Clean_3.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
