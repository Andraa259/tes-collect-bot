import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from datetime import datetime
import io

def get_indo_date():
    months = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 
              7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
    now = datetime.now()
    return f"{now.day} {months[now.month]} {now.year}"

def process_text_replacement(paragraph, expert_name, img_file, ttd_width):
    # 1. Ganti Tanggal & Kunci Jarak Bawah 1cm
    if "Surabaya," in paragraph.text:
        paragraph.text = f"Surabaya, {get_indo_date()}"
        # Set jarak setelah baris ini tepat 1cm
        paragraph.paragraph_format.space_after = Cm(1.0)
        # Pastikan tidak ada spasi tambahan sebelum baris ini
        paragraph.paragraph_format.space_before = Pt(0)
        # Set line spacing ke single agar tidak ada spasi antar baris ekstra
        paragraph.paragraph_format.line_spacing = 1.0

    # 2. Ganti TTD dan Nama & Buat Jarak Atas 0
    if "(Tt Expert Judgement)" in paragraph.text:
        paragraph.text = "" 
        # Pastikan paragraf TTD ini menempel tepat di bawah spasi 1cm tadi
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        
        run = paragraph.add_run()
        run.add_picture(img_file, width=ttd_width)
        run.add_break() 
        run.add_text(f"{expert_name}")

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
