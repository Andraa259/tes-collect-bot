import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from datetime import datetime
import io

# Fungsi untuk mendapatkan tanggal Indonesia
def get_indo_date():
    months = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 
        5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 
        9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    now = datetime.now()
    return f"{now.day} {months[now.month]} {now.year}"

st.title("Auto-Sign Expert Judgement (Final)")

uploaded_file = st.file_uploader("Upload File Word", type=["docx"])
img_file = st.file_uploader("Upload TTD (Pastikan sudah di-crop mepet)", type=["png", "jpg", "jpeg"])
expert_name = st.text_input("Nama Expert")

if st.button("Generate Dokumen"):
    if uploaded_file and img_file and expert_name:
        doc = Document(uploaded_file)
        
        # Ukuran TTD Statis
        TTD_WIDTH = Cm(4.5)
        # Jarak yang kamu mau (0,5 cm)
        TARGET_SPACING = Cm(0.5)

        # 1. Kumpulkan semua paragraf (baik di body utama maupun di dalam tabel)
        all_paragraphs = list(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_paragraphs.extend(cell.paragraphs)

        # 2. Proses Penggantian & Pengaturan Spasi
        for i, p in enumerate(all_paragraphs):
            # Handle Baris Surabaya
            if "Surabaya," in p.text:
                p.text = f"Surabaya, {get_indo_date()}"
                fmt = p.paragraph_format
                fmt.space_before = Pt(0)
                fmt.space_after = TARGET_SPACING # Set ke 0.5 cm
                fmt.line_spacing = 1.0
                
                # Cek jika paragraf setelahnya kosong (biasanya bekas 'Enter' manual), kita hilangkan
                if i + 1 < len(all_paragraphs):
                    next_p = all_paragraphs[i+1]
                    if next_p.text.strip() == "" or "(Tt Expert Judgement)" not in next_p.text:
                        # Jika paragraf setelah Surabaya kosong, kecilkan ukurannya jadi 0
                        if next_p.text.strip() == "":
                            next_p.paragraph_format.space_before = Pt(0)
                            next_p.paragraph_format.space_after = Pt(0)
                            next_p.paragraph_format.line_spacing = Pt(1)

            # Handle Baris TTD & Nama
            if "(Tt Expert Judgement)" in p.text:
                p.text = "" 
                fmt = p.paragraph_format
                fmt.space_before = Pt(0) # Menempel tepat setelah spasi 0.5cm dari atas
                fmt.space_after = Pt(0)
                fmt.line_spacing = 1.0
                
                run = p.add_run()
                run.add_picture(img_file, width=TTD_WIDTH)
                run.add_break() 
                run.add_text(f"{expert_name}")

        # 3. Output
        target_stream = io.BytesIO()
        doc.save(target_stream)
        
        st.success("Dokumen Berhasil Diproses dengan Jarak 0.5 cm!")
        st.download_button(
            label="Download Hasil Final",
            data=target_stream.getvalue(),
            file_name=f"Validated_{expert_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning("Pastikan File, TTD, dan Nama sudah diisi.")
