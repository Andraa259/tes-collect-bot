import streamlit as st
import pandas as pd
import io
import re
import zipfile
from pypdf import PdfReader, PdfWriter

# Konfigurasi halaman
st.set_page_config(page_title="PDF Splitter & Renamer", page_icon="✂️", layout="centered")

st.title("✂️ Certificate PDF Splitter & Renamer")
st.write("Sistem otomatis memotong PDF dari Canva dan menamainya berdasarkan kolom 'Nama' mulai baris ke-2 Excel.")

# Fungsi untuk membersihkan nama file dari karakter ilegal OS (\ / : * ? " < > |)
def clean_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", filename).strip()

# 1. Area Upload File
col1, col2 = st.columns(2)
with col1:
    uploaded_pdf = st.file_uploader("Unggah PDF Gabungan (dari Canva)", type=["pdf"])
with col2:
    uploaded_excel = st.file_uploader("Unggah Excel Data Nama", type=["xlsx", "csv"])

if uploaded_pdf is not None and uploaded_excel is not None:
    try:
        # 2. Baca File Excel / CSV dengan header di baris ke-2 (index 1)
        # KODE YANG BENAR (Baru):
        if uploaded_excel.name.endswith('.csv'):
            df = pd.read_csv(uploaded_excel)
        else:
            df = pd.read_excel(uploaded_excel)
        # Bersihkan baris yang kolom 'Nama'-nya kosong
        if 'Nama' in df.columns:
            df = df.dropna(subset=['Nama']).reset_index(drop=True)
        
        # 3. Baca Struktur File PDF
        pdf_reader = PdfReader(uploaded_pdf)
        total_pages = len(pdf_reader.pages)
        total_rows = len(df)
        
        st.markdown("---")
        st.subheader("📊 Validasi Data")
        st.write(f"📄 Jumlah Halaman PDF: **{total_pages} halaman**")
        st.write(f"📋 Jumlah Baris Data (Kolom Nama): **{total_rows} orang**")
        
        # Validasi ketersediaan kolom 'Nama'
        if 'Nama' not in df.columns:
            st.error("❌ Kolom dengan judul 'Nama' tidak ditemukan di baris ke-2 Excel. Silakan periksa kembali file Anda.")
        else:
            # Validasi kesesuaian jumlah halaman dan data excel
            if total_pages != total_rows:
                st.warning(f"⚠️ **Perhatian:** Jumlah halaman PDF ({total_pages}) tidak sama dengan jumlah baris data Excel ({total_rows}). Pastikan urutannya sudah benar.")
            else:
                st.success("✅ Struktur data cocok! Siap diproses langsung ke file ZIP.")
                
            # 4. Tombol Utama Langsung Ekstrak dan Bungkus ke ZIP
            if st.button("🚀 Mulai Potong & Bungkus ke ZIP", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Mempersiapkan ZIP di dalam memori
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    # Loop sekuensial berdasarkan urutan baris
                    for i in range(min(total_pages, total_rows)):
                        # Mengambil data dari kolom 'Nama' secara absolut
                        raw_name = str(df.iloc[i]['Nama'])
                        formatted_name = raw_name.title() # Format Capitalize Each Word
                        clean_name = clean_filename(formatted_name)
                        
                        # Set nama file PDF individu
                        filename = f"Sertifikat_{clean_name}.pdf"
                        
                        # Proses pemotongan per halaman PDF
                        pdf_writer = PdfWriter()
                        pdf_writer.add_page(pdf_reader.pages[i])
                        
                        # Simpan berkas halaman ke bentuk bytes
                        page_buffer = io.BytesIO()
                        pdf_writer.write(page_buffer)
                        page_bytes = page_buffer.getvalue()
                        
                        # Langsung masukkan berkas ke dalam ZIP
                        zip_file.writestr(filename, page_bytes)
                        
                        # Update visual progress bar di Streamlit UI
                        progress = (i + 1) / min(total_pages, total_rows)
                        progress_bar.progress(progress)
                        status_text.text(f"Mengonversi {i+1}/{min(total_pages, total_rows)}: {filename}")
                
                st.success("🎉 Pemrosesan Selesai! Semua sertifikat telah berhasil dipisahkan.")
                
                # 5. Tombol Unduh File ZIP Akhir
                st.download_button(
                    label="📥 Download Semua Sertifikat (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Sertifikat_Terpisah_Wonderful_Class.zip",
                    mime="application/zip"
                )
                
    except Exception as e:
        st.error(f"Terjadi kesalahan teknis saat memproses berkas: {e}")
