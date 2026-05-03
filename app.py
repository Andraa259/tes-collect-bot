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

st.set_page_config(page_title="Engineer Injector v5", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #00ff41; }
    .stButton>button { border: 1px solid #00ff41 !important; color: #00ff41 !important; background: transparent !important; width: 100%; height: 3.5em; }
    header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# --- CORE FUNCTIONS ---

def get_gsheet_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(creds)

def fetch_all_sheets():
    try:
        client = get_gsheet_client()
        ss = client.open_by_url(GSHEET_URL)
        results = {}
        # Tambahkan 'Pekerjaan' di awal list kolom untuk kebutuhan Word
        cols = ["Nama", "Pekerjaan"] + [f"A{i+1}" for i in range(36)]
        
        for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
            ws = ss.worksheet(s_name)
            data = ws.get("B4:AL33") # GSheets cuma punya Nama + 36 Aitem
            
            # Buat DataFrame
            temp_df = pd.DataFrame(data, columns=["Nama"] + [f"A{i+1}" for i in range(36)])
            # Sisipkan kolom 'Pekerjaan' kosong di index 1 (setelah Nama)
            temp_df.insert(1, "Pekerjaan", "")
            results[s_name] = temp_df
        return results
    except Exception as e:
        st.error(f"GSheets Fetch Error: {e}")
        return None

def send_tele(chat_id, file_buf, fname, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    file_buf.seek(0)
    requests.post(url, data={'chat_id': chat_id, 'caption': caption}, files={'document': (fname, file_buf)})

def generate_word_v5(name, job, scores_kj, scores_rel, scores_kes, template_path):
    doc = Document(template_path)
    for p in doc.paragraphs:
        if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
        if "Pekerjaan\t:" in p.text: p.text = f"Pekerjaan\t: {job}" # Suntik Pekerjaan ke Word
    
    if doc.tables:
        table = doc.tables[0]
        for i in range(36):
            if i + 1 < len(table.rows):
                row = table.rows[i + 1]
                row.cells[3].text, row.cells[4].text, row.cells[5].text = str(scores_kj[i]), str(scores_rel[i]), str(scores_kes[i])
    
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- INTERFACE ---

st.title("🖥️ BATCH INJECTOR V5: IDENTITY SPLIT LOGIC")
st.caption("Pekerjaan diinput di sini untuk Word, tapi tidak akan dikirim ke GSheets.")

if 'db' not in st.session_state:
    st.session_state.db = None

if st.button("📥 FETCH DATABASE"):
    with st.spinner("Syncing..."):
        st.session_state.db = fetch_all_sheets()
        if st.session_state.db: st.success("Data Loaded.")

if st.session_state.db:
    tab_kj, tab_rel, tab_kes = st.tabs(["📊 KEJELASAN (Master Identitas)", "📈 RELEVANSI", "📉 KESESUAIAN"])
    
    with tab_kj:
        st.info("Input Nama & Pekerjaan di Tab ini. Pekerjaan hanya untuk Word.")
        df_kj = st.data_editor(st.session_state.db["KEJELASAN"], num_rows="dynamic", key="e_kj", use_container_width=True)
    with tab_rel:
        df_rel = st.data_editor(st.session_state.db["RELEVANSI"], num_rows="dynamic", key="e_rel", use_container_width=True)
    with tab_kes:
        df_kes = st.data_editor(st.session_state.db["KESESUAIAN"], num_rows="dynamic", key="e_kes", use_container_width=True)

    if st.button("🚀 EXECUTE: SMART SYNC & DUAL TELE"):
        try:
            word_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
            excel_tmpl = "CVI Aiken Zuyy.xlsx"
            orig_len = len(st.session_state.db["KEJELASAN"])
            new_count = len(df_kj) - orig_len
            
            if new_count <= 0:
                st.warning("No new entries.")
            else:
                with st.spinner("Running Pipeline..."):
                    client = get_gsheet_client()
                    ss = client.open_by_url(GSHEET_URL)
                    
                    # --- LOGIC 1: SYNC KE GSHEETS (BUANG KOLOM PEKERJAAN) ---
                    map_dfs = {"KEJELASAN": df_kj, "RELEVANSI": df_rel, "KESESUAIAN": df_kes}
                    for s_name, df_target in map_dfs.items():
                        ws = ss.worksheet(s_name)
                        # Slicing: Ambil kolom 'Nama' (index 0) dan 'A1' s/d 'A36' (index 2 dst)
                        # Kolom 'Pekerjaan' (index 1) ditinggalkan
                        df_to_gsheet = df_target.drop(columns=["Pekerjaan"])
                        ws.update(f"B4:AL{3+len(df_to_gsheet)}", df_to_gsheet.fillna("").values.tolist())

                    # --- LOGIC 2: SEND TELEGRAM (PAKAI PEKERJAAN) ---
                    for i in range(new_count):
                        idx = orig_len + i
                        name = str(df_kj.iloc[idx]["Nama"])
                        job = str(df_kj.iloc[idx]["Pekerjaan"]) # Ambil Pekerjaan buat Word
                        
                        if not name or name == "nan": continue
                        
                        # Generate Word dengan Nama + Job + 3 Skor
                        word_buf = generate_word_v5(
                            name, job, 
                            df_kj.iloc[idx, 2:].tolist(), 
                            df_rel.iloc[idx, 2:].tolist(), 
                            df_kes.iloc[idx, 2:].tolist(), 
                            word_tmpl
                        )
                        
                        send_tele(ID_USER_WORD, word_buf, f"Form_{name}.docx", f"✅ Manual: {name}")
                        word_buf.seek(0)
                        send_tele(ID_USER_FULL, word_buf, f"Form_{name}.docx", f"✅ Log: {name}")
                    
                    st.success(f"Execution Successful. {new_count} records processed.")
                    st.session_state.db = {"KEJELASAN": df_kj, "RELEVANSI": df_rel, "KESESUAIAN": df_kes}
        
        except Exception as e:
            st.error(f"CRASH: {e}")
