import streamlit as st
import pandas as pd
from docx import Document
import io

# --- 1. KONFIGURASI PEKERJAAN (ARRAY URUT) ---
# Sesuaikan isi teks di dalam kurung siku sesuai urutan panelis di Excel
DAFTAR_PEKERJAAN = [
    "Dosen Psikologi",              # Untuk Panelis 1: Juli
    "Mahasiswi Psikologi",          # Untuk Panelis 2: Amelia Nayla
    "Akademisi",                    # Untuk Panelis 3: Farah Tazqia
    "Praktisi Psikologi",           # Untuk Panelis 4: nilon dharu
    "Dosen Psikologi",              # Untuk Panelis 5: Mukhammad Elvino
    "Peneliti Psikometri",          # Untuk Panelis 6: Flora Frederica
    "Mahasiswi Psikologi"           # Untuk Panelis 7: Jeniffer
]

def proses_batch_word(file_excel, template_word):
    sheets = {'KEJELASAN': 'kj', 'RELEVANSI': 'rel', 'KESESUAIAN': 'kes'}
    combined_data = {}

    try:
        # --- 2. EKSTRAKSI DATA DARI EXCEL ---
        for sheet_name, key_code in sheets.items():
            # Membaca data mulai baris ke-4 (index 3)
            df = pd.read_excel(file_excel, sheet_name=sheet_name, header=None)
            
            # Ambil data panelis (baris 4 sampai selesai)
            for row_idx in range(3, len(df)):
                nama_raw = str(df.iloc[row_idx, 1]).strip()
                if nama_raw == "nan" or not nama_raw: 
                    continue
                
                if nama_raw not in combined_data:
                    combined_data[nama_raw] = {}
                
                # Ambil 36 skor (index 2 sampai 37)
                scores = df.iloc[row_idx, 2:38].tolist()
                combined_data[nama_raw][key_code] = scores

        st.divider()
        st.subheader(f"✅ Terdeteksi {len(combined_data)} Panelis")

        # --- 3. GENERATE WORD UNTUK SETIAP PANELIS ---
        list_nama_panelis = list(combined_data.keys())

        for i, nama in enumerate(list_nama_panelis):
            # Reset template untuk setiap orang
            template_word.seek(0)
            doc = Document(template_word)
            
            # Mapping Pekerjaan dari Array
            pekerjaan_aktif = DAFTAR_PEKERJAAN[i] if i < len(DAFTAR_PEKERJAAN) else "Expert Judgment"

            # Isi Nama dan Pekerjaan di Paragraf
            for p in doc.paragraphs:
                if "Nama" in p.text and ":" in p.text:
                    p.text = f"Nama\t\t: {nama}"
                if "Pekerjaan" in p.text and ":" in p.text:
                    p.text = f"Pekerjaan\t: {pekerjaan_aktif}"

            # Isi Tabel Skor
            table = doc.tables[0]
            item_counter = 0
            
            for row in table.rows:
                # Logika Pembersihan Spasi (Resolved)
                text_kolom_aitem = "".join(row.cells[2].text.split()).lower()
                
                # Cek apakah baris ini berisi aitem (Favorable/Unfavorable)
                if "(favorable)" in text_kolom_aitem or "(unfavorable)" in text_kolom_aitem:
                    data_skor = combined_data[nama]
                    if item_counter < len(data_skor['kj']):
                        # Isi kolom KJ, REL, KES (Cell index 3, 4, 5)
                        row.cells[3].text = str(int(float(data_skor['kj'][item_counter])))
                        row.cells[4].text = str(int(float(data_skor['rel'][item_counter])))
                        row.cells[5].text = str(int(float(data_skor['kes'][item_counter])))
                        item_counter += 1

            # --- 4. OUTPUT DOWNLOAD BUTTON ---
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            
            st.download_button(
                label=f"📥 Download Form: {nama}",
                data=buf,
                file_name=f"Form_Validasi_{nama.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"btn_{i}"
            )

    except Exception as e:
        st.error(f"Terjadi kesalahan teknis: {e}")

# --- UI UTAMA ---
st.set_page_config(page_title="Final Word Batch Generator", layout="centered")
st.title("📄 Forgiveness Scale: Batch Word Generator")
st.markdown("Sistem ini akan menulis ulang 7 file Word secara otomatis berdasarkan data Excel.")

with st.expander("ℹ️ Instruksi"):
    st.write("1. Upload file Excel CVI Aiken (Kumulatif).")
    st.write("2. Upload template Word (Form yang kolom skornya masih kosong).")
    st.write("3. Klik tombol 'Proses' untuk memunculkan link download setiap panelis.")

col1, col2 = st.columns(2)
with col1:
    up_excel = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
with col2:
    up_word = st.file_uploader("Upload Template Word (.docx)", type=["docx"])

if up_excel and up_word:
    if st.button("🚀 PROSES & REWRITE SEMUA DATA", use_container_width=True):
        proses_batch_word(up_excel, up_word)
