import streamlit as st
import pandas as pd
from docx import Document
import openpyxl
import gspread
from google.oauth2.service_account import Credentials
import io
import requests
import time

# --- ENGINEER CONFIG ---
TOKEN = st.secrets["TOKEN"]
ID_USER_WORD = st.secrets.get("CHAT_ID_1")
ID_USER_FULL = st.secrets.get("CHAT_ID_2")
GSHEET_URL = st.secrets["GSHEET_URL"]

st.set_page_config(page_title="Engineer Injector Final", layout="wide")

# Theme Programmer: Dark, Sleek, & High Contrast
st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #00ff41; }
    .stButton>button { border: 1px solid #00ff41 !important; color: #00ff41 !important; background: transparent !important; width: 100%; height: 3.5em; font-family: 'Courier New', monospace; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { background-color: #1e1e1e; border-radius: 5px; color: white; padding: 10px; border: 1px solid #333; }
    .stTabs [aria-selected="true"] { border: 1px solid #00ff41 !important; color: #00ff41 !important; }
    .stFileUploader { border: 1px dashed #00ff41 !important; border-radius: 10px; padding: 10px; }
    header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# --- CORE FUNCTIONS ---

def get_gsheet_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(creds)

def fetch_all_sheets():
    """Tarik data dari GSheets (Range B4:AL33)"""
    try:
        client = get_gsheet_client()
        ss = client.open_by_url(GSHEET_URL)
        results = {}
        # Kolom Nama (1) + Kolom Skor (36) = 37 kolom dari GSheets
        cols_gs = ["Nama"] + [f"A{i+1}" for i in range(36)]
        
        for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
            ws = ss.worksheet(s_name)
            data = ws.get("B4:AL33")
            df = pd.DataFrame(data, columns=cols_gs)
            # Sisipkan kolom Pekerjaan kosong buat diisi di Streamlit/Excel Upload
            df.insert(1, "Pekerjaan", "")
            results[s_name] = df
        return results
    except Exception as e:
        st.error(f"GSheets Fetch Error: {e}")
        return None

def send_tele(chat_id, file_buf, fname, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    file_buf.seek(0)
    requests.post(url, data={'chat_id': chat_id, 'caption': caption}, files={'document': (fname, file_buf)})

def generate_word_final(name, job, scores_kj, scores_rel, scores_kes, template_path):
    """Injeksi data ke template Word"""
    doc = Document(template_path)
    for p in doc.paragraphs:
        if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
        if "Pekerjaan\t:" in p.text: p.text = f"Pekerjaan\t: {job}"
    
    if doc.tables:
        table = doc.tables[0]
        for i in range(36):
            if i + 1 < len(table.rows):
                row = table.rows[i + 1]
                # Kolom 3: KJ, 4: REL, 5: KES
                row.cells[3].text = str(scores_kj[i])
                row.cells[4].text = str(scores_rel[i])
                row.cells[5].text = str(scores_kes[i])
    
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- INTERFACE ---

st.title("🖥️ ENGINEER BATCH INJECTOR FINAL (v7)")
st.code(f"MAPPING: A(Nama), B(Pekerjaan), C-AL(36 Items)")

if 'db' not in st.session_state:
    st.session_state.db = None

# Sidebar Controls
with st.sidebar:
    st.header("📡 GSheets Connection")
    if st.button("📥 FETCH DATA PUSAT"):
        with st.spinner("Linking Database..."):
            st.session_state.db = fetch_all_sheets()
            if st.session_state.db: st.success("Connected.")

# File Uploader (Solusi Android)
st.subheader("📂 Step 1: Upload Master File (Excel)")
up_file = st.file_uploader("Upload .xlsx dengan 3 Sheet (KEJELASAN, RELEVANSI, KESESUAIAN)", type=["xlsx"])

if up_file and st.session_state.db:
    if st.button("⚙️ PROCESS & MERGE UPLOAD"):
        try:
            xl = pd.ExcelFile(up_file)
            for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
                if s_name in xl.sheet_names:
                    new_data = xl.parse(s_name, header=None) # Header=None biar lo bisa langsung isi di baris 1
                    new_data.columns = ["Nama", "Pekerjaan"] + [f"A{i+1}" for i in range(36)]
                    # Gabungkan dengan data fetch (biar baris bertambah di bawahnya)
                    st.session_state.db[s_name] = pd.concat([st.session_state.db[s_name], new_data], ignore_index=True)
            st.success("File Merged. Cek tabel di bawah.")
        except Exception as e:
            st.error(f"Error Processing Excel: {e}")

# Data Editor Area
if st.session_state.db:
    st.subheader("📝 Step 2: Review & Edit Data")
    t_kj, t_rel, t_kes = st.tabs(["[KJ] KEJELASAN", "[REL] RELEVANSI", "[KES] KESESUAIAN"])
    
    with t_kj:
        df_kj = st.data_editor(st.session_state.db["KEJELASAN"], num_rows="dynamic", key="final_kj", use_container_width=True)
    with t_rel:
        df_rel = st.data_editor(st.session_state.db["RELEVANSI"], num_rows="dynamic", key="final_rel", use_container_width=True)
    with t_kes:
        df_kes = st.data_editor(st.session_state.db["KESESUAIAN"], num_rows="dynamic", key="final_kes", use_container_width=True)

    # Execution Area
    st.write("---")
    st.subheader("🚀 Step 3: Execution")
    if st.button("⚡ START INJECTION PIPELINE"):
        try:
            w_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
            orig_len = len(st.session_state.db["KEJELASAN"])
            new_count = len(df_kj) - orig_len
            
            if new_count <= 0:
                st.warning("Gak ada data baru yang dideteksi.")
            else:
                with st.spinner("Processing..."):
                    client = get_gsheet_client()
                    ss = client.open_by_url(GSHEET_URL)
                    
                    # 1. Sync ke GSheets (Kolom Pekerjaan di-drop otomatis)
                    map_sync = {"KEJELASAN": df_kj, "RELEVANSI": df_rel, "KESESUAIAN": df_kes}
                    for s_name, df_target in map_sync.items():
                        ws = ss.worksheet(s_name)
                        df_to_gs = df_target.drop(columns=["Pekerjaan"]) # Kolom 'Pekerjaan' tidak dikirim ke GSheets
                        ws.update(f"B4:AL{3+len(df_to_gs)}", df_to_gs.fillna("").values.tolist())

                    # 2. Injeksi Word & Kirim Telegram (Hanya untuk baris BARU)
                    for i in range(new_count):
                        idx = orig_len + i
                        name = str(df_kj.iloc[idx]["Nama"])
                        job = str(df_kj.iloc[idx]["Pekerjaan"])
                        
                        if not name or name == "nan": continue
                        
                        # Ambil skor dari 3 tab berbeda di baris yang sama
                        s_kj = df_kj.iloc[idx, 2:].tolist()
                        s_rel = df_rel.iloc[idx, 2:].tolist()
                        s_kes = df_kes.iloc[idx, 2:].tolist()
                        
                        w_buf = generate_word_final(name, job, s_kj, s_rel, s_kes, w_tmpl)
                        
                        # Kirim ke User 1 (Word Only)
                        send_tele(ID_USER_WORD, w_buf, f"Form_{name}.docx", f"✅ Manual Input: {name}")
                        
                        # Kirim ke User 2 (Word Only - Excel kumulatif bisa ditarik manual dari GSheets)
                        w_buf.seek(0)
                        send_tele(ID_USER_FULL, w_buf, f"Form_{name}.docx", f"✅ Engineer Log: {name}")
                        
                    st.success(f"Pipeline Finished. {new_count} data panelis sukses diproses.")
                    st.session_state.db = {"KEJELASAN": df_kj, "RELEVANSI": df_rel, "KESESUAIAN": df_kes}
        except Exception as e:
            st.error(f"PIPELINE_CRASH: {e}")
else:
    st.info("Klik 'FETCH DATA PUSAT' di sidebar untuk menyalakan engine.")
