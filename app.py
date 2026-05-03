import streamlit as st
import pandas as pd
from docx import Document
import io
import requests

# --- 1. KONFIGURASI TELEGRAM & PEKERJAAN ---
# Pastikan TOKEN dan ID sudah ada di Streamlit Secrets
TOKEN = st.secrets["TOKEN"]
# Masukkan daftar Chat ID tujuan di sini (bisa satu atau lebih)
TARGET_CHAT_IDS = [st.secrets["CHAT_ID_1"], st.secrets.get("CHAT_ID_2")] 

DAFTAR_PEKERJAAN = [
    "Dosen Psikologi",              # Juli
    "Mahasiswi Psikologi",          # Amelia Nayla
    "Akademisi",                    # Farah Tazqia
    "Praktisi Psikologi",           # nilon dharu
    "Dosen Psikologi",              # Mukhammad Elvino
    "Peneliti Psikometri",          # Flora Frederica
    "Mahasiswi Psikologi"           # Jeniffer
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
            'caption': f"✅ **Batch Report**: {panelis_name}\nStatus: Berhasil di-generate otomatis."
        }
        try:
            response = requests.post(url, data=payload, files=files)
            if response.status_code == 200:
                status_list.append(True)
            else:
                status_list.append(False)
        except:
            status_list.append(False)
    return all(status_list)

# --- 3. LOGIKA CORE BATCH PROCESS ---
def proses_batch_dan_kirim(file_excel, template_word):
    sheets = {'KEJELASAN': 'kj', 'RELEVANSI': 'rel', 'KESESUAIAN': 'kes'}
    combined_data = {}

    try:
        # Ekstraksi Data Excel
        for sheet_name, key_code in sheets.items():
            df = pd.read_excel(file_excel, sheet_name=sheet_name, header=None)
            for row_idx in range(3, len(df)):
                nama_raw = str(df.iloc[row_idx, 1]).strip()
                if nama_raw == "nan" or not nama_raw: continue
                if nama_raw not in combined_data: combined_data[nama_raw] = {}
                scores = df.iloc[row_idx, 2:38].tolist()
                combined_data[nama_raw][key_code] = scores

        list_nama = list(combined_data.keys())
        st.write(f"🔄 Memproses {len(list_nama)} panelis...")

        for i, nama in enumerate(list_nama):
            template_word.seek(0)
            doc = Document(template_word)
            pekerjaan_aktif = DAFTAR_PEKERJAAN[i] if i < len(DAFTAR_PEKERJAAN) else "Expert Judgment"

            # Fill Header
            for p in doc.paragraphs:
                if "Nama" in p.text and ":" in p.text: p.text = f"Nama\t\t: {nama}"
                if "Pekerjaan" in p.text and ":" in p.text: p.text = f"Pekerjaan\t: {pekerjaan_aktif}"

            # Fill Table
            table = doc.tables[0]
            item_counter = 0
            for row in table.rows:
                text_clean = "".join(row.cells[2].text.split()).lower()
                if "(favorable)" in text_clean or "(unfavorable)" in text_clean:
                    d = combined_data[nama]
                    if item_counter < len(d['kj']):
                        row.cells[3].text = str(int(float(d['kj'][item_counter])))
                        row.cells[4].text = str(int(float(d['rel'][item_counter])))
                        row.cells[5].text = str(int(float(d['kes'][item_counter])))
                        item_counter += 1

            # Simpan dan Kirim
            word_buf = io.BytesIO()
            doc.save(word_buf)
            fname = f"Form_Validasi_{nama.replace(' ', '_')}.docx"
            
            with st.status(f"Mengirim file {nama}...", expanded=False) as status:
                if kirim_ke_telegram(word_buf, fname, nama):
                    status.update(label=f"✅ {nama} terkirim ke Telegram!", state="complete")
                else:
                    status.update(label=f"❌ {nama} gagal dikirim.", state="error")

        st.balloons()
        st.success("Semua data telah diproses dan dikirim ke Telegram!")

    except Exception as e:
        st.error(f"Error: {e}")

# --- 4. UI STREAMLIT ---
st.set_page_config(page_title="Tele-Batch Generator", page_icon="🤖")
st.title("🤖 Forgiveness Bot: Auto-Tele Sender")
st.markdown("Upload file untuk mengirim hasil rewrite langsung ke Telegram tim.")

up_excel = st.file_uploader("Upload Excel Kumulatif", type=["xlsx"])
up_word = st.file_uploader("Upload Template Word", type=["docx"])

if up_excel and up_word:
    if st.button("🚀 KIRIM SEMUA KE TELEGRAM", use_container_width=True):
        proses_batch_dan_kirim(up_excel, up_word)
