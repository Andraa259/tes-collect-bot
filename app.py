import streamlit as st
from docx import Document
from docx.shared import Cm
from datetime import datetime
import io

def get_indo_date():
    months = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 
              7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
    now = datetime.now()
    return f"{now.day} {months[now.month]} {now.year}"

def process_text_replacement(paragraph, expert_name, img_file, ttd_width):
    # 1. Ganti Tanggal (Cari kata Surabaya saja agar lebih fleksibel)
    if "Surabaya," in paragraph.text:
        # Menghapus garis bawah dan mengganti dengan tanggal baru
        # Kita pakai split untuk menjaga kata 'Surabaya,' tetap ada
        paragraph.text = f"Surabaya, {get_indo_date()}"

    # 2. Ganti TTD dan Nama
    if "(Tt Expert Judgement)" in paragraph.text:
        paragraph.text = "" # Hapus teks aslinya
        run = paragraph.add_run()
        run.add_picture(img_file, width=ttd_width) # Masukkan foto TTD
        run.add_break() 
        run.add_text(f"({expert_name})") # Masukkan Nama

st.title("Auto-Sign Fixer")

uploaded_file = st.file_uploader("Upload Word", type=["docx"])
img_file = st.file_uploader("Upload TTD", type=["png", "jpg"])
expert_name = st.text_input("Nama Expert")

if st.button("Proses Sekarang"):
    if uploaded_file and img_file and expert_name:
        doc = Document(uploaded_file)
        ttd_size = Cm(4.5)

        # PROSES PARAGRAF BIASA
        for p in doc.paragraphs:
            process_text_replacement(p, expert_name, img_file, ttd_size)

        # PROSES DALAM TABEL (Ini kunci perbaikannya!)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        process_text_replacement(p, expert_name, img_file, ttd_size)

        target_stream = io.BytesIO()
        doc.save(target_stream)
        st.success("Selesai! Silakan download.")
        st.download_button("Download Hasil", target_stream.getvalue(), "Hasil_Final.docx")
