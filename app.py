import streamlit as st
import pandas as pd
from docx import Document
import io
import requests

# --- 1. KONFIGURASI TELEGRAM & DATA ARRAY ---
TOKEN = st.secrets["TOKEN"]
TARGET_CHAT_IDS = [st.secrets["CHAT_ID_1"], st.secrets.get("CHAT_ID_2")] 

# Array Pekerjaan (Urut sesuai Excel)
DAFTAR_PEKERJAAN = [
    "Mahasiswa",                                                                       # Juli
    "Mahasiswa Psikologi Semester 4 di Universitas 17 Agustus 1945 Surabaya",          # Amelia Nayla
    "-",                                                                               # Farah Tazqia
    "mahasiswa",                                                                       # nilon dharu
    "Mahasiswa",                                                                       # Mukhammad Elvino
    "Mahasiswa",                                                                       # Flora Frederica
    "Mahasiswi"                                                                        # Jeniffer
]

# Array Catatan Akhir (Hanya Amelia dan Nilon yang ada isinya)
DAFTAR_CATATAN = [
    "",                                      # Juli (Kosong)
    "Item yang dibuat sudah cukup baik dan sudah sesuai dengan indikator yang ada.",  # Amelia Nayla (Ada isi)
    "",                                      # Farah Tazqia (Kosong)
    "cukup jelas dan baik",   # nilon dharu (Ada isi)
    "",                                      # Mukhammad Elvino (Kosong)
    "",                                      # Flora Frederica (Kosong)
    ""                                       # Jeniffer (Kosong)
]

# --- 2. FUNGSI PENGIRIMAN TELEGRAM ---
def kirim_ke_telegram(file_buf, file_name, panelis_name):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    status_list = []
    
    for chat_id in TARGET_CHAT_IDS:
        if not chat_id: continue
        file_buf.seek(0)
        files = {'document': (file_name, file_buf)}
        payload = {
            'chat_id': chat_id,
            'caption': f"✅ Batch Report: {panelis_name}\nStatus: Berhasil di-generate otomatis."
        }
        try:
            response = requests.post(url, data=payload, files=files)
            status_list.append(response.status_code == 200)
        except:
            status_list.append(False)
    return all(status_list)

# --- 3. HELPER UNTUK MENGATASI NaN ---
def format_skor(val):
    if pd.isna(val) or val == "":
        return "0"
    try:
        return str(int(float(val)))
    except:
        return "0"

# --- 4. LOGIKA CORE BATCH PROCESS ---
def proses_batch_dan_kirim(file_excel, template_word):
    sheets = {'KEJELASAN': 'kj', 'RELEVANSI': 'rel', 'KESESUAIAN': 'kes'}
    combined_data = {}

    try:
        # Ekstraksi Data Excel
        for sheet_name, key_code in sheets.items():
            df = pd.read_excel(file_excel, sheet_name=sheet_name, header=None)
            for row_idx in range(3, 10):
                nama_raw = str(df.iloc[row_idx, 1]).strip()
                if nama_raw == "nan" or not nama_raw or nama_raw == "None": continue
                if nama_raw not in combined_data: combined_data[nama_raw] = {}
                scores = df.iloc[row_idx, 2:38].tolist()
                combined_data[nama_raw][key_code] = scores

        list_nama = list(combined_data.keys())[:7]
        st.write(f"🔄 Memproses {len(list_nama)} panelis...")

        for i, nama in enumerate(list_nama):
            template_word.seek(0)
            doc = Document(template_word)
            
            # Ambil Pekerjaan & Catatan dari Array berdasarkan index i
            pekerjaan_aktif = DAFTAR_PEKERJAAN[i] if i < len(DAFTAR_PEKERJAAN) else "Expert Judgment"
            catatan_aktif = DAFTAR_CATATAN[i] if i < len(DAFTAR_CATATAN) else ""

            # Fill Header (Nama & Pekerjaan)
            for p in doc.paragraphs:
                if "Nama" in p.text and ":" in p.text: p.text = f"Nama\t\t: {nama}"
                if "Pekerjaan" in p.text and ":" in p.text: p.text = f"Pekerjaan\t: {pekerjaan_aktif}"

            # Fill Table (Skor & Catatan Akhir)
            table = doc.tables[0]
            item_counter = 0
            for row in table.rows:
                text_raw = row.cells[2].text
                text_clean = "".join(text_raw.split()).lower()
                
                # Input Skor Aitem
                if "(favorable)" in text_clean or "(unfavorable)" in text_clean:
                    d = combined_data[nama]
                    if item_counter < len(d['kj']):
                        row.cells[3].text = format_skor(d['kj'][item_counter])
                        row.cells[4].text = format_skor(d['rel'][item_counter])
                        row.cells[5].text = format_skor(d['kes'][item_counter])
                        item_counter += 1
                
                # Input Catatan Akhir (Jika baris mengandung kata 'Catatan')
                if "catatan" in text_raw.lower() and catatan_aktif:
                    # Menambahkan catatan ke sel yang sama (biasanya kolom keterangan/catatan)
                    row.cells[2].text = f"{text_raw}\n{catatan_aktif}"

            # Simpan dan Kirim
            word_buf = io.BytesIO()
            doc.save(word_buf)
            fname = f"Form Validasi Expert Judgement Forgiveness_{nama.replace(' ', '_')}.docx"
            
            with st.status(f"Mengirim file {nama}...", expanded=False) as status:
                if kirim_ke_telegram(word_buf, fname, nama):
                    status.update(label=f"✅ {nama} terkirim!", state="complete")
                else:
                    status.update(label=f"❌ {nama} gagal.", state="error")

        st.balloons()
        st.success("Batch pengiriman selesai!")

    except Exception as e:
        st.error(f"Error: {e}")

# --- 5. UI STREAMLIT ---
st.set_page_config(page_title="Tele-Batch Generator", page_icon="🤖")
st.title("🤖 Forgiveness Bot: Auto-Sender + Global Notes")

up_excel = st.file_uploader("Upload Excel Kumulatif", type=["xlsx"])
up_word = st.file_uploader("Upload Template Word", type=["docx"])

if up_excel and up_word:
    if st.button("🚀 PROSES & KIRIM KE TELEGRAM", use_container_width=True):
        proses_batch_dan_kirim(up_excel, up_word)
