import streamlit as st
import pandas as pd
from docx import Document
import gspread
from google.oauth2.service_account import Credentials
import io
import requests
import time

# --- 1. CONFIG & SECRETS ---
st.set_page_config(page_title="Engineer Master Hybrid v7.8", layout="wide")

TOKEN = st.secrets["TOKEN"]
ID_USER_WORD = st.secrets.get("CHAT_ID_1") 
ID_USER_FULL = st.secrets.get("CHAT_ID_2")
GSHEET_URL = st.secrets["GSHEET_URL"]

# --- 2. CORE FUNCTIONS ---

def get_gsheet_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(creds)

def find_first_empty_row(ws):
    """Cek antrean kosong berdasarkan Kolom C (Skor Pertama)"""
    col_c = ws.col_values(3) 
    for r in range(4, 34): 
        if r > len(col_c) or not col_c[r-1]:
            return r
    return 34

def generate_word_final(name, job, skj, srel, skes, template_path):
    """Injeksi data Nama, Pekerjaan, dan 36 baris Skor ke Word"""
    try:
        doc = Document(template_path)
        # 1. Injeksi Identitas
        for p in doc.paragraphs:
            if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
            if "Pekerjaan\t:" in p.text: p.text = f"Pekerjaan\t: {job}"
        
        # 2. Injeksi Tabel Skor (Asumsi Tabel 0, Baris 1-36, Kolom 3,4,5)
        if doc.tables:
            table = doc.tables[0]
            for i in range(36):
                # Baris i+1 karena baris 0 biasanya header tabel
                if i + 1 < len(table.rows):
                    row = table.rows[i + 1]
                    # Kolom 3: Kejelasan, 4: Relevansi, 5: Kesesuaian
                    row.cells[3].text = str(skj[i])
                    row.cells[4].text = str(srel[i])
                    row.cells[5].text = str(skes[i])
        
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf
    except Exception as e:
        st.error(f"Word Error: {e}")
        return None

def send_tele(chat_id, file_buf, fname, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    file_buf.seek(0)
    requests.post(url, data={'chat_id': chat_id, 'caption': caption}, files={'document': (fname, file_buf)})

# --- 3. UI INTERFACE ---

st.title("🖥️ MASTER HYBRID INJECTOR v7.8")
st.markdown("---")

if 'batch_data' not in st.session_state:
    st.session_state.batch_data = None

# STEP 1: UPLOAD & PARSING
up_file = st.file_uploader("Upload Excel Aiken", type=["xlsx"])

if up_file:
    if st.button("⚙️ PROCESS & ALIGN DATA"):
        try:
            xl = pd.ExcelFile(up_file)
            temp_db = {}
            for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
                # Baca Excel, paksa kolom 38 (A-AL)
                raw_xl = pd.read_excel(up_file, sheet_name=s_name, header=None, skiprows=3)
                raw_xl = raw_xl.reindex(columns=range(38), fill_value=0)
                
                rows = []
                for i in range(len(raw_xl)):
                    name = raw_xl.iloc[i, 1] # Kolom B
                    if pd.isna(name) or str(name).strip() == "": continue
                    
                    row_dict = {"Nama": str(name), "Pekerjaan": ""}
                    for a_idx in range(36):
                        val = raw_xl.iloc[i, 2 + a_idx] # Kolom C dst
                        row_dict[f"A{a_idx+1}"] = val if not pd.isna(val) else 0
                    rows.append(row_dict)
                temp_db[s_name] = pd.DataFrame(rows)
            
            st.session_state.batch_data = temp_db
            st.success("Data Excel terbaca rapi. Silakan cek tabel di bawah.")
            st.rerun()
        except Exception as e:
            st.error(f"Parsing Error: {e}")

# STEP 2: REVIEW & INJECT
if st.session_state.batch_data:
    tabs = st.tabs(["📊 KJ", "📈 REL", "📉 KES"])
    sheets = ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]
    for i, tab in enumerate(tabs):
        with tab:
            st.session_state.batch_data[sheets[i]] = st.data_editor(
                st.session_state.batch_data[sheets[i]], 
                num_rows="dynamic", 
                key=f"editor_{sheets[i]}"
            )

    st.write("---")
    if st.button("🚀 EXECUTE SMART SYNC (WORD + GSHEETS)"):
        try:
            with st.spinner("Antrekan data & Injecting Word..."):
                client = get_gsheet_client()
                ss = client.open_by_url(GSHEET_URL)
                w_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
                
                db_kj = st.session_state.batch_data["KEJELASAN"]
                db_rel = st.session_state.batch_data["RELEVANSI"]
                db_kes = st.session_state.batch_data["KESESUAIAN"]

                for idx in range(len(db_kj)):
                    name = db_kj.iloc[idx]["Nama"]
                    job = db_kj.iloc[idx]["Pekerjaan"]
                    
                    # 1. Update GSheets (Smart Append)
                    for s_name in sheets:
                        ws = ss.worksheet(s_name)
                        target_row = find_first_empty_row(ws)
                        
                        if target_row <= 33:
                            # Kirim Nama & 36 Skor
                            skor = st.session_state.batch_data[s_name].iloc[idx, 2:].tolist()
                            ws.update_cell(target_row, 2, name)
                            cells = ws.range(target_row, 3, target_row, 3 + len(skor) - 1)
                            for s_idx, val in enumerate(skor): cells[s_idx].value = val
                            ws.update_cells(cells)

                    # 2. Generate Word & Telegram
                    # Ambil semua kategori skor untuk orang yang sama
                    skj_list = db_kj.iloc[idx, 2:].tolist()
                    srel_list = db_rel.iloc[idx, 2:].tolist()
                    skes_list = db_kes.iloc[idx, 2:].tolist()
                    
                    word_file = generate_word_final(name, job, skj_list, srel_list, skes_list, w_tmpl)
                    
                    if word_file:
                        send_tele(ID_USER_WORD, word_file, f"Form_{name}.docx", f"✅ Word: {name}")
                        word_file.seek(0)
                        send_tele(ID_USER_FULL, word_file, f"Form_{name}.docx", f"✅ Full Log: {name}")
                
                st.success("Semua data berhasil disinkronkan!")
                st.session_state.batch_data = None
        except Exception as e:
            st.error(f"Pipeline Error: {e}")
