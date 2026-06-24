import streamlit as st
import pandas as pd
import io
import re
import zipfile
from pypdf import PdfReader, PdfWriter

# Konfigurasi halaman
st.set_page_config(page_title="PDF Splitter - Organized ZIP", page_icon="✂️", layout="centered")

st.title("✂️ Certificate PDF Splitter & Renamer (Panitia & Peserta)")
st.write("Sistem otomatis memotong PDF Canva dan membaginya ke dalam folder **Panitia** (Baris 1-28) dan **Peserta** (Baris 29-54) di dalam file ZIP.")

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
        # 2. Baca File Excel / CSV (Judul di Baris 1, Data mulai Baris 2)
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
            st.error("❌ Kolom dengan judul 'Nama' tidak ditemukan. Pastikan tulisan judul kolom di baris pertama bernama 'Nama'.")
        else:
            # Validasi kesesuaian jumlah halaman dan data excel
            if total_pages != total_rows:
                st.warning(f"⚠️ **Perhatian:** Jumlah halaman PDF ({total_pages}) tidak sama dengan jumlah baris data Excel ({total_rows}). Pastikan urutan file sudah benar.")
            else:
                st.success("✅ Struktur data cocok! Siap diproses ke dalam folder ZIP.")
                
            # 4. Tombol Utama Proses Ekstrak dan Kelompokkan Folder
            if st.button("🚀 Mulai Potong & Kelompokkan ke ZIP", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Mempersiapkan ZIP di dalam memori
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    # Loop sekuensial berdasarkan urutan baris data
                    for i in range(min(total_pages, total_rows)):
                        baris_data_ke = i + 1  # Urutan baris (1-indexed)
                        
                        # LOGIKA PERCABANGAN DISTRIBUSI FOLDER:
                        if 1 <= baris_data_ke <= 28:
                            folder_name = "Panitia"
                        elif 29 <= baris_data_ke <= 54:
                            folder_name = "Peserta"
                        else:
                            folder_name = "Lainnya" # Antisipasi jika ada data melebihi baris 54
                        
                        # Mengambil data nama untuk baris saat ini
                        raw_name = str(df.iloc[i]['Nama'])
                        formatted_name = raw_name.title() # Format Capitalize Each Word
                        clean_name = clean_filename(formatted_name)
                        
                        # Struktur path di dalam ZIP: Nama_Folder/Sertifikat_Nama.pdf
                        filename_in_zip = f"{folder_name}/{clean_name}.pdf"
                        
                        # Proses pemotongan halaman PDF
                        pdf_writer = PdfWriter()
                        pdf_writer.add_page(pdf_reader.pages[i])
                        
                        # Simpan berkas ke bentuk bytes
                        page_buffer = io.BytesIO()
                        pdf_writer.write(page_buffer)
                        page_bytes = page_buffer.getvalue()
                        
                        # Masukkan ke dalam ZIP sesuai dengan struktur foldernya
                        zip_file.writestr(filename_in_zip, page_bytes)
                        
                        # Update progress bar UI
                        progress = (i + 1) / min(total_pages, total_rows)
                        progress_bar.progress(progress)
                        status_text.text(f"📁 [{folder_name}] Mengonversi {i+1}: Sertifikat_{clean_name}.pdf")
                
                st.success("🎉 Pemrosesan Selesai! Semua sertifikat telah dipisahkan ke dalam folder masing-masing.")
                
                # 5. Tombol Unduh File ZIP Akhir
                st.download_button(
                    label="📥 Download Semua Sertifikat (ZIP Organized)",
                    data=zip_buffer.getvalue(),
                    file_name="Sertifikat_Wonderful_Class_Organized.zip",
                    mime="application/zip"
                )
                
    except Exception as e:
        st.error(f"Terjadi kesalahan teknis saat memproses berkas: {e}")
