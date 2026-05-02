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

st.title("Auto-Sign Expert (Zero Ghost Space)")

uploaded_file = st.file_uploader("Upload Word", type=["docx"])
img_file = st.file_uploader("Upload TTD", type=["png", "jpg", "jpeg"])
expert_name = st.text_input("Nama Expert")

if st.button("Generate Dokumen"):
    if uploaded_file and img_file and expert_name:
        doc = Document(uploaded_file)
        TTD_WIDTH = Cm(4.5)

        # 1. Kumpulkan semua paragraf dari body dan tabel
        all_paragraphs = list(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_paragraphs.extend(cell.paragraphs)

        # 2. Proses Manipulasi
        for p in all_paragraphs:
            # LOGIKA A: Gabungkan Tanggal, TTD, dan Nama di baris 'Surabaya'
            if "Surabaya," in p.text:
                # Bersihkan paragraf agar tidak ada formatting lama yang mengganggu
                p.text = "" 
                run = p.add_run(f"Surabaya, {get_indo_date()}")
                
                # Enter sekali sebelum TTD
                run.add_break() 
                run.add_picture(img_file, width=TTD_WIDTH)
                
                # Enter sekali sebelum Nama
                run.add_break() 
                run.add_text(f"({expert_name})")
                
                # Pengaturan spasi agar mepet
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.space_before = Pt(0)

            # LOGIKA B: Hapus total paragraf '(Tt Expert Judgement)'
            if "(Tt Expert Judgement)" in p.text:
                # Menghapus elemen paragraf secara permanen dari dokumen
                p_element = p._element
                p_element.getparent().remove(p_element)

        # 3. Download
        target_stream = io.BytesIO()
        doc.save(target_stream)
        
        st.success("Selesai! Spasi siluman telah dihapus.")
        st.download_button(
            label="Download Hasil Final",
            data=target_stream.getvalue(),
            file_name=f"Validated_Fixed_{expert_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
