import streamlit as st
import pandas as pd
from io import BytesIO

def process_expert_mapping_v4(file):
    # Membaca file Excel
    df = pd.read_excel(file)
    
    # Indeks Nama Panelis adalah 3
    name_col_idx = 3
    name_col_name = df.columns[name_col_idx]
    
    # 1. Start dari baris 3 (Jennifer)
    # Dalam pandas, jika header adalah baris 0, maka baris 2 Excel adalah index 0.
    # Baris 3 Excel (Jennifer) adalah index 1.
    df_start = df.iloc[1:].copy()
    
    # 2. Algoritma Mapping per 4 Kolom (Mulai Indeks 5)
    # Urutan: i (Kejelasan), i+1 (Relevansi), i+2 (Kesesuaian), i+3 (Keterangan - Skip)
    kej_indices = []
    rel_indices = []
    kes_indices = []
    
    # Melakukan loop di seluruh kolom mulai dari indeks 5 dengan langkah 4
    for i in range(5, len(df.columns), 4):
        if i < len(df.columns):
            kej_indices.append(i)
        if i + 1 < len(df.columns):
            rel_indices.append(i + 1)
        if i + 2 < len(df.columns):
            kes_indices.append(i + 2)

    def extract_clean_data(indices):
        # Mengambil kolom Nama + Kolom Skor terpilih
        return df_start.iloc[:, [name_col_idx] + indices]

    # 3. Export ke Excel dengan 3 Sheet
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        extract_clean_data(kej_indices).to_excel(writer, sheet_name='Kejelasan', index=False)
        extract_clean_data(rel_indices).to_excel(writer, sheet_name='Relevansi', index=False)
        extract_clean_data(kes_indices).to_excel(writer, sheet_name='Kesesuaian', index=False)
        
    return output.getvalue(), len(kej_indices)

# --- UI Streamlit ---
st.set_page_config(page_title="Expert Judgement Mapper V4", layout="wide")
st.title("🧩 Expert Mapping: Per 4 Kolom (Start Jennifer)")
st.write("Memproses file **Formulir tanpa judul (Jawaban) (1)_2.xlsx** dengan aturan pemetaan posisi.")

uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

if uploaded_file:
    with st.spinner("Memetakan aitem..."):
        try:
            processed_file, total_aitem = process_expert_mapping_v4(uploaded_file)
            
            st.success(f"Mapping Selesai! Berhasil memetakan **{total_aitem} aitem**.")
            st.info("Data dimulai dari Jennifer. Kolom 'Keterangan' dan panelis 'Zaki' telah dihapus.")
            
            st.download_button(
                label="📥 Download Data Bersih",
                data=processed_file,
                file_name="Data_Expert_Mapped_Jennifer.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
