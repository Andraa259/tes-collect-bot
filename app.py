import streamlit as st
import pandas as pd
from docx import Document
import gspread
from google.oauth2.service_account import Credentials
import io
import requests
import time

# --- 1. CONFIG ---
st.set_page_config(page_title="Engineer Hybrid Injector v7.7", layout="wide")

# --- 2. CREDENTIALS ---
TOKEN = st.secrets["TOKEN"]
ID_USER_WORD = st.secrets.get("CHAT_ID_1") 
ID_USER_FULL = st.secrets.get("CHAT_ID_2")
GSHEET_URL = st.secrets["GSHEET_URL"]

# --- 3. CORE FUNCTIONS ---

def get_gsheet_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(creds)

def find_first_empty_row(ws):
    """Logika yang sama dengan Portal lo: cari baris kosong berdasarkan Kolom C (Skor)"""
    col_c = ws.col_values(3) # Kolom C adalah skor pertama
    for r in range(4, 34): # Range baris 4 sampai 33
        if r > len(col_c) or not col_c[r-1]:
            return r
    return 34 # Jika penuh

def generate_word_final(name, job, skj, srel, skes, template_path):
    doc = Document(template_path)
    for p in doc.paragraphs:
        if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
        if "Pekerjaan\t:" in p.text: p.text = f"Pekerjaan\t: {job}"
    if doc.tables:
        table = doc.tables[0]
        for i in range(36):
            if i + 1 < len(table.rows):
                row = table.rows[i + 1]
                row.cells[3].text, row.cells[4].text, row.cells[5].text = str(skj[i]), str(srel[i]), str(skes[i])
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

def send_tele(chat_id, file_buf, fname, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    file_buf.seek(0)
    requests.post(url, data={'chat_id': chat_id, 'caption': caption}, files={'document': (fname, file_buf)})

# --- 4. INTERFACE ---

st.title("🖥️ HYBRID BATCH INJECTOR v7.7")
st.info("Sistem ini akan mencari baris kosong di GSheets agar tidak menimpa data dari Web Portal.")

if 'batch_data' not in st.session_state:
    st.session_state.batch_data = None

# Step 1: Upload Excel
up_file = st.file_uploader("Upload Excel (Format Aiken)", type=["xlsx"])

if up_file:
    if st.button("⚙️ PROCESS EXCEL"):
        try:
            xl = pd.ExcelFile(up_file)
            temp_db = {}
            for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
                raw_xl = pd.read_excel(up_file, sheet_name=s_name, header=None, skiprows=3)
                raw_xl = raw_xl.reindex(columns=range(38), fill_value=0)
                
                rows = []
                for i in range(len(raw_xl)):
                    name = raw_xl.iloc[i, 1]
                    if pd.isna(name) or str(name).strip() == "": continue
                    
                    row_dict = {"Nama": str(name), "Pekerjaan": ""}
                    for a_idx in range(36):
                        val = raw_xl.iloc[i, 2 + a_idx]
                        row_dict[f"A{a_idx+1}"] = val if not pd.isna(val) else 0
                    rows.append(row_dict)
                temp_db[s_name] = pd.DataFrame(rows)
            st.session_state.batch_data = temp_db
            st.success("Excel siap diantrekan!")
        except Exception as e:
            st.error(f"Error: {e}")

# Step 2: Review & Execute
if st.session_state.batch_data:
    tabs = st.tabs(["📊 KJ", "📈 REL", "📉 KES"])
    sheets = ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]
    for i, tab in enumerate(tabs):
        with tab:
            st.session_state.batch_data[sheets[i]] = st.data_editor(st.session_state.batch_data[sheets[i]], num_rows="dynamic", key=f"ed_{sheets[i]}")

    if st.button("🚀 EXECUTE SMART APPEND"):
        try:
            with st.spinner("Mencari antrean kosong di GSheets..."):
                client = get_gsheet_client()
                ss = client.open_by_url(GSHEET_URL)
                w_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
                
                db_kj = st.session_state.batch_data["KEJELASAN"]
                db_rel = st.session_state.batch_data["RELEVANSI"]
                db_kes = st.session_state.batch_data["KESESUAIAN"]

                for idx in range(len(db_kj)):
                    name = db_kj.iloc[idx]["Nama"]
                    job = db_kj.iloc[idx]["Pekerjaan"]
                    
                    # --- LOGIKA SINKRONISASI ---
                    # Cari baris kosong di setiap sheet untuk setiap orang
                    for s_name in sheets:
                        ws = ss.worksheet(s_name)
                        target_row = find_first_empty_row(ws) # <--- SAMA DENGAN PORTAL
                        
                        if target_row > 33:
                            st.warning(f"GSheets {s_name} penuh! Tidak bisa input {name}")
                            continue
                        
                        # Data yang akan dikirim (Nama + 36 Skor)
                        skor = st.session_state.batch_data[s_name].iloc[idx, 2:].tolist()
                        ws.update_cell(target_row, 2, name) # Kolom B
                        
                        # Batch update skor secara horizontal mulai kolom C (3)
                        cells = ws.range(target_row, 3, target_row, 3 + len(skor) - 1)
                        for s_idx, val in enumerate(skor):
                            cells[s_idx].value = val
                        ws.update_cells(cells)

                    # --- KIRIM TELEGRAM ---
                    skj = db_kj.iloc[idx, 2:].tolist()
                    srel = db_rel.iloc[idx, 2:].tolist()
                    skes = db_kes.iloc[idx, 2:].tolist()
                    
                    w_buf = generate_word_final(name, job, skj, srel, skes, w_tmpl)
                    send_tele(ID_USER_WORD, w_buf, f"Form_{name}.docx", f"✅ {name}")
                    w_buf.seek(0); send_tele(ID_USER_FULL, w_buf, f"Form_{name}.docx", f"✅ Log: {name}")
                
                st.success("Batching selesai tanpa menimpa data User Web!")
                st.session_state.batch_data = None
        except Exception as e:
            st.error(f"Pipeline Error: {e}")
