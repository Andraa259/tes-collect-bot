import streamlit as st
import pandas as pd
from io import BytesIO

def process_expert_mapping(file):
    # Membaca file secara utuh
    df = pd.read_excel(file)
    
    # Identifikasi kolom nama panelis
    name_col_idx = 3 # Indeks kolom 'Nama Panelis'
    name_col_name = df.columns[name_col_idx]
    
    # 1. Hapus baris dengan nama 'zaki' (Case-insensitive)
    # Ini otomatis membuat Jennifer (baris 3) menjadi data pertama
    df = df[df[name_col_name].str.contains('zaki', case=False, na=False) == False].copy()
    
    # 2. Algoritma Mapping per 4 Kolom
    # Start dari indeks 5 (Aitem pertama), melompat setiap 4 kolom
    kejelasan_indices = []
    relevansi_indices = []
    kesesuaian_indices = []
    
    for i in range(5, len(df.columns), 4):
        # Pastikan tidak out of index
        if i < len(df.columns):
            kejelasan_indices.append(i)
        if i + 1 < len(df.columns):
            relevansi_indices.append(i + 1)
        if i + 2 < len(df.columns):
            kesesuaian_indices.append(i + 2)
        # Indeks i + 3 (Keterangan) dilewati/dibuang sesuai perintah
            
    # 3. Fungsi untuk mengambil Nama + Kolom Skor saja
    def get_mapped_df(indices):
        # Gabungkan indeks nama panelis dengan indeks skor yang dipilih
        selected_indices = [name_col_idx] + indices
        return df.iloc[:, selected_indices]

    # 4. Proses Simpan ke Excel dengan 3 Sheet
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        get_mapped_df(kejelasan_indices).to_excel(writer, sheet_name='Kejelasan', index=False)
        get_mapped_df(relevansi_indices).to_excel(writer, sheet_name='Relevansi', index=False)
        get_mapped_df(kesesuaian_indices).to_excel(writer, sheet_name='Kesesuaian', index=False)
        
    return output.getvalue(), len(kejelasan_indices)

# --- UI Streamlit ---
st.set_page_config(page_title="Expert Mapping Tool", layout="wide")
st.title("🧩 Expert Judgement Positional Mapping")
st.write("Algoritma: Mapping otomatis setiap 4 kolom mulai dari kolom ke-6 (indeks 5).")

uploaded_file = st.file_uploader("Upload 'Formulir tanpa judul (Jawaban) (1).xlsx'", type=["xlsx"])

if uploaded_file:
    with st.spinner("Memproses mapping kolom..."):
        try:
            processed_file, total_aitem = process_expert_mapping(uploaded_file)
            
            st.success(f"Berhasil! Terdeteksi **{total_aitem} aitem** melalui algoritma per 4 kolom.")
            st.info("Baris 'Zaki' telah dihapus. Data bersih dimulai dari Jennifer.")
            
            st.download_button(
                label="📥 Download Excel (Nama & Skor Saja)",
                data=processed_file,
                file_name="Hasil_Mapping_Expert.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
