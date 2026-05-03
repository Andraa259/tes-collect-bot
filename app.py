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

st.set_page_config(page_title="Engineer Injector v4", layout="wide")

# Dark Theme Programmer
st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #00ff41; }
    .stButton>button { border: 1px solid #00ff41 !important; color: #00ff41 !important; background: transparent !important; width: 100%; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { background-color: #1e1e1e; border-radius: 5px; color: white; padding: 10px; }
    .stTabs [aria-selected="true"] { border: 1px solid #00ff41 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- CORE LOGIC ---

def get_gsheet_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(creds)

def fetch_all_sheets():
    """Tarik 3 sheet sekaligus agar data sinkron"""
    client = get_gsheet_client()
    ss = client.open_by_url(GSHEET_URL)
    results = {}
    cols = ["Nama"] + [f"A{i+1}" for i in range(36)]
    
    for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
        ws = ss.worksheet(s_name)
        data = ws.get("B4:AL33")
        results[s_name] = pd.DataFrame(data, columns=cols)
    return results

def generate_word_v4(name, scores_kj, scores_rel, scores_kes, template_path):
    """Suntik 3 jenis skor berbeda ke dalam satu baris tabel Word"""
    doc = Document(template_path)
    for p in doc.paragraphs:
        if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
    
    if doc.tables:
        table = doc.tables[0]
        for i in range(36): # 36 aitem
            if i + 1 < len(table.rows):
                row = table.rows[i + 1]
                # Kolom 3: Kejelasan | Kolom 4: Relevansi | Kolom 5: Kesesuaian
                row.cells[3].text = str(scores_kj[i])
                row.cells[4].text = str(scores_rel[i])
                row.cells[5].text = str(scores_kes[i])
    
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- UI INTERFACE ---

st.title("🖥️ BATCH INJECTOR V4: 3-SHEET SYNC")

if 'db' not in st.session_state:
    st.session_state.db = None

if st.button("📥 FETCH 3-SHEET DATABASE"):
    with st.spinner("Synchronizing..."):
        st.session_state.db = fetch_all_sheets()
        st.success("All Sheets Loaded.")

if st.session_state.db:
    # Buat Tab untuk masing-masing kriteria
    tab_kj, tab_rel, tab_kes = st.tabs(["📊 KEJELASAN", "📈 RELEVANSI", "📉 KESESUAIAN"])
    
    with tab_kj:
        df_kj = st.data_editor(st.session_state.db["KEJELASAN"], num_rows="dynamic", key="edit_kj", use_container_width=True)
    with tab_rel:
        df_rel = st.data_editor(st.session_state.db["RELEVANSI"], num_rows="dynamic", key="edit_rel", use_container_width=True)
    with tab_kes:
        df_kes = st.data_editor(st.session_state.db["KESESUAIAN"], num_rows="dynamic", key="edit_kes", use_container_width=True)

    st.write("---")
    if st.button("🚀 EXECUTE: SYNC ALL SHEETS & SEND DUAL TELE"):
        try:
            word_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
            excel_tmpl = "CVI Aiken Zuyy.xlsx"
            
            # Deteksi data baru (asumsi jumlah baris baru sama di ketiga tab)
            orig_len = len(st.session_state.db["KEJELASAN"])
            new_entries_kj = df_kj.iloc[orig_len:]
            
            if new_entries_kj.empty:
                st.warning("No new entries to process.")
            else:
                with st.spinner("Processing Hybrid Pipeline..."):
                    # 1. Update GSheets (3 Sheet Sekaligus)
                    client = get_gsheet_client()
                    ss = client.open_by_url(GSHEET_URL)
                    map_dfs = {"KEJELASAN": df_kj, "RELEVANSI": df_rel, "KESESUAIAN": df_kes}
                    
                    for s_name, df_target in map_dfs.items():
                        ws = ss.worksheet(s_name)
                        ws.update(f"B4:AL{3+len(df_target)}", df_target.fillna("").values.tolist())

                    # 2. Generate & Send (Word disuntik 3 skor sekaligus)
                    for i in range(len(new_entries_kj)):
                        row_kj = df_kj.iloc[orig_len + i]
                        row_rel = df_rel.iloc[orig_len + i]
                        row_kes = df_kes.iloc[orig_len + i]
                        
                        name = str(row_kj["Nama"])
                        if not name or name == "nan": continue
                        
                        # Generate Word dengan 3 sumber skor
                        word_buf = generate_word_v4(
                            name, 
                            row_kj[1:].tolist(), 
                            row_rel[1:].tolist(), 
                            row_kes[1:].tolist(), 
                            word_tmpl
                        )
                        
                        # Kirim Tele
                        # User 1
                        send_tele(ID_USER_WORD, word_buf, f"Form_{name}.docx", f"✅ Manual Input: {name}")
                        # User 2 + Excel
                        word_buf.seek(0)
                        send_tele(ID_USER_FULL, word_buf, f"Form_{name}.docx", f"✅ Engineer Log: {name}")
                        # Logic Excel Master Kumulatif (Fungsi sync_excel_master dari kode sebelumnya)
                        # ...
                
                st.success("All systems green. 3 Sheets synced & Tele sent.")
                st.session_state.db = {"KEJELASAN": df_kj, "RELEVANSI": df_rel, "KESESUAIAN": df_kes}

else:
    st.info("Initiate 'FETCH' to manage data.")
