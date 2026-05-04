import streamlit as st
import pandas as pd
from io import BytesIO

def process_expert_mapping_final(file):
    # Membaca file Excel
    df = pd.read_excel(file)
    
    # Indeks Nama Panelis tetap di 3
    name_col_idx = 3
    
    # Algoritma Mapping per 4 Kolom (Mulai Indeks 5)
    # Urutan: Kejelasan (i), Relevansi (i+1), Kesesuaian (i+2), Keterangan (i+3)
    kej_indices = []
    rel_indices = []
    kes_indices = []
    
    # Loop melalui seluruh kolom mulai dari indeks 5 dengan langkah 4
    for i in range(5, len(df.columns), 4):
        if i < len(df.columns):
            kej_indices.append(i)
        if i + 1 < len(df.columns):
            rel_indices.append(i + 1)
        if i + 2 < len(df.columns):
            kes_indices.append(i + 2)
        # Indeks i + 3 (Keterangan) otomatis terlewati

    def create_clean_sheet(indices):
        # Mengambil kolom Nama + Kolom Skor berdasarkan indeks posisi
        selected = [name_col_idx] + indices
        return df.iloc[:, selected]

    # Export ke Excel dengan 3 Sheet berbeda
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        create_clean_sheet(kej_indices).to_excel(writer, sheet_name='Kejelasan', index=False)
        create_clean_sheet(rel_indices).to_excel(writer, sheet_name='Relevansi', index=False)
        create_clean_sheet(kes_indices).to_excel(writer, sheet_name='Kesesuaian', index=False)
        
    return output.getvalue(), len(kej_indices)

# --- Antarmuka Streamlit ---
st.set_page_config(page_title="Expert Mapping Final", layout="wide")
st.title("🧩 Expert Judgement Mapper (Positional Algorithm)")
st.write("Memproses **seluruh baris data** dengan pemetaan setiap 4 kolom.")

uploaded_file = st.file_uploader("Upload file 'Form Validasi Expert judgement (Jawaban)_2.xlsx'", type=["xlsx"])

if uploaded_file:
    with st.spinner("Sedang memproses seluruh baris..."):
        try:
            processed_file, total_aitem = process_expert_mapping_final(uploaded_file)
            
            st.success(f"Mapping Selesai! Berhasil memetakan **{total_aitem} aitem** dari seluruh baris input.")
            
            st.download_button(
                label="📥 Download Excel Hasil Mapping",
                data=processed_file,
                file_name="Data_Expert_Judgement_Mapped.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan teknis: {e}")
