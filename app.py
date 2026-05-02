import streamlit as st
from docx import Document
from docx.shared import Cm
from datetime import datetime
import io

# Fungsi untuk konversi nama bulan ke Bahasa Indonesia
def get_indo_date():
    months = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 
        5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 
        9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    now = datetime.now()
    return f"{now.day} {months[now.month]} {now.year}"

st.title("Auto-Sign & Fill Form")
st.write("Aplikasi ini akan mengisi tanggal, nama, dan TTD pada file 'Form Validasi Expert Judgement'.")

# 1. Input Section
uploaded_file = st.file_uploader("Upload File Word (Form Validasi)", type=["docx"])
img_file = st.file_uploader("Upload Foto TTD Online", type=["png", "jpg", "jpeg"])
expert_name = st.text_input("Masukkan Nama Expert (untuk mengganti (Tt Expert Judgement))")

# Konfigurasi Ukuran TTD (Terkunci/Statis)
TTD_WIDTH = Cm(4.5)  # Anda bisa ubah ukuran ini sesuai kebutuhan

if st.button("Proses Dokumen"):
    if uploaded_file and img_file and expert_name:
        doc = Document(uploaded_file)
        today_date = get_indo_date()
        
        # 2. Proses Pencarian dan Penggantian
        for paragraph in doc.paragraphs:
            # Ganti Tanggal
            if "Surabaya, ___________________" in paragraph.text:
                paragraph.text = paragraph.text.replace("Surabaya, ___________________", f"Surabaya, {today_date}")
            
            # Ganti TTD dan Nama
            if "(Tt Expert Judgement)" in paragraph.text:
                # Bersihkan teks lama
                paragraph.text = ""
                run = paragraph.add_run()
                
                # Tambahkan TTD (Ukuran otomatis terkunci di sini)
                run.add_picture(img_file, width=TTD_WIDTH)
                
                # Tambahkan Nama di bawah TTD
                run.add_break() # Memberi jarak antara gambar dan nama
                run.add_text(f"({expert_name})")
                
                # Opsional: Membuat teks menjadi rata tengah (jika diperlukan)
                # paragraph.alignment = 1 

        # 3. Download Section
        target_stream = io.BytesIO()
        doc.save(target_stream)
        
        st.success("Dokumen berhasil diproses!")
        st.download_button(
            label="Download File Hasil",
            data=target_stream.getvalue(),
            file_name=f"Validated_Form_{expert_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("Mohon lengkapi semua input (File, TTD, dan Nama).")
