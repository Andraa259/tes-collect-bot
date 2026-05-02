import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from datetime import datetime
import io

def get_indo_date():
    months = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 
        5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 
        9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    now = datetime.now()
    return f"{now.day} {months[now.month]} {now.year}"

st.title("Auto-Sign Expert (One-Block Logic)")

uploaded_file = st.file_uploader("Upload Word", type=["docx"])
img_file = st.file_uploader("Upload TTD", type=["png", "jpg", "jpeg"])
expert_name = st.text_input("Nama Expert")

if st.button("Generate Dokumen"):
    if uploaded_file and img_file and expert_name:
        doc = Document(uploaded_file)
        TTD_WIDTH = Cm(4.5)

        # 1. Ambil semua paragraf (body + tabel)
        all_paragraphs = list(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_paragraphs.extend(cell.paragraphs)

        # 2. Eksekusi Logika Baru
        for p in all_paragraphs:
            # Cari baris Surabaya
            if "Surabaya," in p.text:
                # Masukkan Tanggal
                p.text = f"Surabaya, {get_indo_date()}"
                
                # Masukkan TTD dan Nama di paragraf yang SAMA agar rapat
                run = p.add_run()
                run.add_break() # Pindah baris untuk TTD
                run.add_picture(img_file, width=TTD_WIDTH)
                run.add_break() # Pindah baris untuk Nama
                run.add_text(f"{expert_name}")
                
                # Paksa spasi jadi single dan nol
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.space_before = Pt(0)

            # Cari baris (Tt Expert Judgement) dan HAPUS total
            if "(Tt Expert Judgement)" in p.text:
                p.text = ""
                # Ciutkan sisa paragrafnya agar tidak memakan ruang
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = Pt(1)

        # 3. Download
        target_stream = io.BytesIO()
        doc.save(target_stream)
        
        st.success("Selesai! Logika gabungan berhasil diterapkan.")
        st.download_button(
            label="Download Hasil",
            data=target_stream.getvalue(),
            file_name=f"Validated_{expert_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
