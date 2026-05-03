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
# Pastikan st.secrets sudah lengkap di Streamlit Cloud
TOKEN = st.secrets["TOKEN"]
ID_USER_WORD = st.secrets.get("CHAT_ID_1")  # User 1: Word Only
ID_USER_FULL = st.secrets.get("CHAT_ID_2")  # User 2: Word & Master Excel
GSHEET_URL = st.secrets["GSHEET_URL"]

st.set_page_config(page_title="Engineer Injector v4", layout="wide")

# Theme Programmer: Dark & Green
st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #00ff41; }
    .stButton>button { border: 1px solid #00ff41 !important; color: #00ff41 !important; background: transparent !important; width: 100%; height: 3.5em; font-family: 'Courier New', monospace; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { background-color: #1e1e1e; border-radius: 5px; color: white; padding: 10px; border: 1px solid #333; }
    .stTabs [aria-selected="true"] { border: 1px solid #00ff41 !important; color: #00ff41 !important; }
    header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# --- CORE FUNCTIONS ---

def get_gsheet_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(creds)

def fetch_all_sheets():
    """Tarik 3 sheet dari GSheets Pusat"""
    try:
        client = get_gsheet_client()
        ss = client.open_by_url(GSHEET_URL)
        results = {}
        cols = ["Nama"] + [f"A{i+1}" for i in range(36)]
        
        for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
            ws = ss.worksheet(s_name)
            data = ws.get("B4:AL33")
            results[s_name] = pd.DataFrame(data, columns=cols)
        return results
    except Exception as e:
        st.error(f"GSheets Fetch Error: {e}")
        return None

def send_tele(chat_id, file_buf, fname, caption):
    """Kirim dokumen ke Telegram API"""
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    file_buf.seek(0)
    requests.post(url, data={'chat_id': chat_id, 'caption': caption}, files={'document': (fname, file_buf)})

def generate_word_v4(name, scores_kj, scores_rel, scores_kes, template_path):
    """Injeksi 3 jenis skor ke satu file Word"""
    doc = Document(template_path)
    for p in doc.paragraphs:
        if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
    
    if doc.tables:
        table = doc.tables[0]
        # Skip header, loop aitem 1-36
        for i in range(36):
            if i + 1 < len(table.rows):
                row = table.rows[i + 1]
                # Kolom 3: KJ | Kolom 4: REL | Kolom 5: KES
                row.cells[3].text = str(scores_kj[i])
                row.cells[4].text = str(scores_rel[i])
                row.cells[5].text = str(scores_kes[i])
    
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def sync_excel_master(template_path, df_kj, df_rel, df_kes):
    """Update file Master Excel kumulatif untuk User 2"""
    wb = openpyxl.load_workbook(template_path)
    map_dfs = {"KEJELASAN": df_kj, "RELEVANSI": df_rel, "KESESUAIAN": df_kes}
    
    for s_name, df_target in map_dfs.items():
        if s_name in wb.sheetnames:
            ws = wb[s_name]
            for idx, row in df_target.iterrows():
                target_r = 4 + idx
                if target_r > 33: break
                ws.cell(row=target_r, column=2, value=row["Nama"])
                for c_idx, val in enumerate(row[1:]):
                    try: ws.cell(row=target_r, column=3 + c_idx, value=int(float(val)))
                    except: ws.cell(row=target_r, column=3 + c_idx, value=0)
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# --- INTERFACE ENGINE ---

st.title("🖥️ ENGINEER BATCH INJECTOR v4")
st.code(f"MASTER_DB: {GSHEET_URL}")

if 'db' not in st.session_state:
    st.session_state.db = None

# 1. FETCH DATA
col_btn, _ = st.columns([1, 2])
with col_btn:
    if st.button("📥 FETCH 3-SHEET DATABASE"):
        with st.spinner("Accessing GSheets..."):
            st.session_state.db = fetch_all_sheets()
            if st.session_state.db:
                st.success("Synchronized.")

# 2. DATA GRID AREA
if st.session_state.db:
    st.subheader("📊 Batch Data Editor (3 Categories)")
    tab_kj, tab_rel, tab_kes = st.tabs(["[KJ] KEJELASAN", "[REL] RELEVANSI", "[KES] KESESUAIAN"])
    
    with tab_kj:
        df_kj = st.data_editor(st.session_state.db["KEJELASAN"], num_rows="dynamic", key="e_kj", use_container_width=True)
    with tab_rel:
        df_rel = st.data_editor(st.session_state.db["RELEVANSI"], num_rows="dynamic", key="e_rel", use_container_width=True)
    with tab_kes:
        df_kes = st.data_editor(st.session_state.db["KESESUAIAN"], num_rows="dynamic", key="e_kes", use_container_width=True)

    # 3. EXECUTION
    st.write("---")
    if st.button("🚀 EXECUTE: SYNC GSHEETS & DUAL TELEGRAM SEND"):
        try:
            # File harus ada di root GitHub
            word_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
            excel_tmpl = "CVI Aiken Zuyy.xlsx"
            
            # Deteksi data baru (Hanya baris yang baru lo paste/tambah)
            orig_len = len(st.session_state.db["KEJELASAN"])
            new_entries_count = len(df_kj) - orig_len
            
            if new_entries_count <= 0:
                st.warning("No new entries detected. Process terminated.")
            else:
                with st.spinner("Executing Pipeline..."):
                    # A. Backup/Overwrite GSheets Pusat (3 Sheets)
                    client = get_gsheet_client()
                    ss = client.open_by_url(GSHEET_URL)
                    map_update = {"KEJELASAN": df_kj, "RELEVANSI": df_rel, "KESESUAIAN": df_kes}
                    
                    for s_name, df_target in map_update.items():
                        ws = ss.worksheet(s_name)
                        # Push all data back to GSheets
                        ws.update(f"B4:AL{3+len(df_target)}", df_target.fillna("").values.tolist())

                    # B. Generate Word & Send Telegram (Hanya untuk baris BARU)
                    for i in range(new_entries_count):
                        idx = orig_len + i
                        row_kj = df_kj.iloc[idx]
                        row_rel = df_rel.iloc[idx]
                        row_kes = df_kes.iloc[idx]
                        
                        name = str(row_kj["Nama"])
                        if not name or name == "nan": continue
                        
                        # Injeksi 3 skor sekaligus ke Word
                        word_buf = generate_word_v4(
                            name, 
                            row_kj[1:].tolist(), 
                            row_rel[1:].tolist(), 
                            row_kes[1:].tolist(), 
                            word_tmpl
                        )
                        
                        # Telegram User 1: Word Only
                        send_tele(ID_USER_WORD, word_buf, f"Form_{name}.docx", f"✅ Manual Input: {name}")
                        
                        # Telegram User 2: Word + Excel Kumulatif
                        word_buf.seek(0)
                        send_tele(ID_USER_FULL, word_buf, f"Form_{name}.docx", f"✅ Log: {name}")
                        
                        excel_buf = sync_excel_master(excel_tmpl, df_kj, df_rel, df_kes)
                        send_tele(ID_USER_FULL, excel_buf, f"Master_Aiken_Update_{name}.xlsx", f"📊 Aiken Report Updated")
                    
                    st.success(f"Execution Successful. {new_entries_count} panelist(s) processed.")
                    # Update state agar tidak dianggap baru di klik berikutnya
                    st.session_state.db = {"KEJELASAN": df_kj, "RELEVANSI": df_rel, "KESESUAIAN": df_kes}
        
        except Exception as e:
            st.error(f"ENGINE_CRASH: {str(e)}")

else:
    st.info("System Ready. Waiting for 'FETCH DATABASE' command.")

st.markdown("---")
st.caption("v4.0 | Engineering Mode | Surabaya, 2026")
