import streamlit as st
import pandas as pd
from io import BytesIO

def process_psp_mapping(file):
    # Membaca file Excel
    df = pd.read_excel(file)
    
    # Indeks Nama Panelis adalah 3 (NAMA)
    name_col_idx = 3
    name_col_name = df.columns[name_col_idx]
    
    # Algoritma Mapping per 3 Kolom (Mulai Indeks 5)
    # i = Kejelasan, i+1 = Relevansi, i+2 = Kesesuaian
    kej_indices = []
    rel_indices = []
    kes_indices = []
    
    # Loop melalui seluruh kolom mulai dari indeks 5 dengan langkah 3
    for i in range(5, len(df.columns), 3):
        if i < len(df.columns):
            kej_indices.append(i)
        if i + 1 < len(df.columns):
            rel_indices.append(i + 1)
        if i + 2 < len(df.columns):
            kes_indices.append(i + 2)

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
st.set_page_config(page_title="PSP Mapping Tool", layout="wide")
st.title("📑 PSP Item Positional Mapper (3-Column Step)")
st.write("Algoritma: Mapping otomatis setiap 3 kolom mulai dari kolom ke-6 (indeks 5).")

uploaded_file = st.file_uploader("Upload file 'Aitem PSP_2.xlsx'", type=["xlsx"])

if uploaded_file:
    with st.spinner("Sedang memproses pemetaan kolom..."):
        try:
            processed_file, total_aitem = process_psp_mapping(uploaded_file)
            
            st.success(f"Mapping Selesai! Berhasil memetakan **{total_aitem} aitem** secara presisi.")
            st.info("Email, Timestamp, dan Universitas telah dibersihkan.")
            
            st.download_button(
                label="📥 Download Excel Hasil Mapping",
                data=processed_file,
                file_name="Aitem_PSP_Cleaned_Mapped.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan teknis: {e}")
