import streamlit as st
import pandas as pd
from io import BytesIO

def process_full_data(file):
    # Membaca seluruh isi file tanpa batasan kolom
    df = pd.read_excel(file)
    
    # Nama kolom panelis sesuai file kamu
    name_col = 'Nama Panelis (Jika memiliki gelar, mohon disertakan)'
    
    # 1. Hapus baris 'zaki' agar data mulai dari Jennifer
    # Kita gunakan case=False agar aman dari perbedaan huruf kapital
    df_filtered = df[df[name_col].str.contains('zaki', case=False, na=False) == False].copy()
    
    # 2. Ambil semua 54 aitem (menggunakan partial match untuk handle typo spasi)
    # Ini akan mencari di seluruh kolom dari A sampai kolom terakhir (jauh melampaui BC)
    kejelasan_cols = [col for col in df.columns if '[Kejelasan' in col]
    relevansi_cols = [col for col in df.columns if '[Relevansi' in col]
    kesesuaian_cols = [col for col in df.columns if '[Kesesuaian' in col]

    def create_sheet(cols):
        # Hanya ambil kolom Nama dan kolom skor yang dipilih
        return df_filtered[[name_col] + cols]

    # 3. Proses simpan ke Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        create_sheet(kejelasan_cols).to_excel(writer, sheet_name='Kejelasan', index=False)
        create_sheet(relevansi_cols).to_excel(writer, sheet_name='Relevansi', index=False)
        create_sheet(kesesuaian_cols).to_excel(writer, sheet_name='Kesesuaian', index=False)
        
    return output.getvalue(), len(kejelasan_cols)

# --- UI ---
st.set_page_config(page_title="Expert Judgement Cleaner", layout="wide")
st.title("📊 Expert Judgement Cleaner (Full 54 Items)")
st.write("Program ini menscan **seluruh kolom** file Anda, bukan hanya sampai BC.")

uploaded_file = st.file_uploader("Upload file 'Formulir tanpa judul (Jawaban) (1).xlsx'", type=["xlsx"])

if uploaded_file:
    with st.spinner("Memproses seluruh aitem..."):
        try:
            processed_file, count = process_full_data(uploaded_file)
            
            # Menampilkan informasi hasil agar kamu tenang
            st.success(f"Berhasil! Menemukan total **{count} aitem** (semua kolom terangkut).")
            st.info("Baris 'Zaki' telah dihapus. Data sekarang dimulai dari Jennifer.")
            
            st.download_button(
                label="📥 Download Hasil Pembersihan",
                data=processed_file,
                file_name="Expert_Judgement_Final_54_Aitem.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")
