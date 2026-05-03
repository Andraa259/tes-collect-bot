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
ID_USER_WORD = st.secrets.get("CHAT_ID_1")  # User 1: Word Only
ID_USER_FULL = st.secrets.get("CHAT_ID_2")  # User 2: Word & Master Excel
GSHEET_URL = st.secrets["GSHEET_URL"]

st.set_page_config(page_title="Engineer Injector", layout="wide")

# Programmer Dark Mode Theme
st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #00ff41; }
    .stButton>button { background-color: #1e1e1e !important; color: #00ff41 !important; border: 1px solid #00ff41 !important; width: 100%; height: 3em; font-family: 'Courier New', monospace; }
    .stDataEditor { border: 1px solid #333 !important; }
    header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# --- CORE LOGIC ---

def get_gsheet_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(creds)

def fetch_gsheets():
    """Tarik database pusat"""
    client = get_gsheet_client()
    ss = client.open_by_url(GSHEET_URL)
    ws = ss.worksheet("KEJELASAN")
    # Tarik kolom Nama (B) dan Skor Aitem (C-AL)
    data = ws.get("B4:AL33")
    cols = ["Nama"] + [f"Aitem_{i+1}" for i in range(36)]
    return pd.DataFrame(data, columns=cols)

def send_tele(chat_id, file_buf, fname, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    file_buf.seek(0)
    requests.post(url, data={'chat_id': chat_id, 'caption': caption}, files={'document': (fname, file_buf)})

def generate_word(name, scores, template_path):
    doc = Document(template_path)
    # Inject Nama
    for p in doc.paragraphs:
        if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
    
    # Inject Tabel
    if doc.tables:
        table = doc.tables[0]
        for i, score in enumerate(scores):
            if i + 1 < len(table.rows):
                row = table.rows[i + 1]
                # Isi Kejelasan, Relevansi, Kesesuaian (Kolom 3, 4, 5)
                for col_idx in [3, 4, 5]:
                    row.cells[col_idx].text = str(score)
    
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def sync_excel_master(template_path, full_df):
    """Update file Excel Aiken secara kumulatif untuk User 2"""
    wb = openpyxl.load_workbook(template_path)
    for sheet_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for idx, row in full_df.iterrows():
                target_r = 4 + idx
                if target_r > 33: break
                ws.cell(row=target_r, column=2, value=row["Nama"])
                for col_idx, val in enumerate(row[1:]): # Aitem scores
                    try: ws.cell(row=target_r, column=3 + col_idx, value=int(float(val)))
                    except: ws.cell(row=target_r, column=3 + col_idx, value=0)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# --- INTERFACE ---

st.title("🖥️ ENGINEER BATCH INJECTOR v3")
st.code(f"Database: {GSHEET_URL}")

# 1. FETCH DATA
if 'engine_df' not in st.session_state:
    st.session_state.engine_df = pd.DataFrame()

col_ctrl, col_empty = st.columns([1, 2])
with col_ctrl:
    if st.button("📥 FETCH DATABASE"):
        with st.spinner("Fetching..."):
            st.session_state.engine_df = fetch_gsheets()
            st.success("Sync Complete.")

# 2. BATCH INPUT AREA
if not st.session_state.engine_df.empty:
    st.subheader("📊 Data Grid (Supports Batch Paste)")
    # Data editor memungkinkan lo Copy-Paste massal dari Excel/Spreadsheet
    edited_df = st.data_editor(
        st.session_state.engine_df,
        num_rows="dynamic",
        use_container_width=True,
        key="main_editor"
    )

    # 3. ACTION: SYNC & TELEGRAM DUAL SEND
    st.write("---")
    if st.button("🚀 EXECUTE: SYNC GSHEETS & DUAL TELEGRAM SEND"):
        # File Template dipanggil dari GitHub Root
        try:
            word_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
            excel_tmpl = "CVI Aiken Zuyy.xlsx"
            
            # Deteksi Baris Baru (Manual entries only)
            original_len = len(st.session_state.engine_df)
            new_entries = edited_df.iloc[original_len:]

            if new_entries.empty:
                st.warning("No new manual entries detected.")
            else:
                with st.spinner("Processing Hybrid Pipeline..."):
                    # A. Backup / Overwrite GSheets (Total Update)
                    client = get_gsheet_client()
                    ss = client.open_by_url(GSHEET_URL)
                    for s in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
                        ws = ss.worksheet(s)
                        ws.update(f"B4:AL{3+len(edited_df)}", edited_df.fillna("").values.tolist())
                    
                    # B. Generate & Send Telegram (Only for NEW entries)
                    for _, row in new_entries.iterrows():
                        name = str(row["Nama"])
                        scores = row[1:].tolist()
                        
                        if not name or name == "nan": continue
                        
                        # Generate Word Local
                        word_buf = generate_word(name, scores, word_tmpl)
                        
                        # USER 1: Word Only
                        send_tele(ID_USER_WORD, word_buf, f"Form_{name}.docx", f"✅ Manual Input: {name}")
                        
                        # USER 2: Word + Master Excel Terupdate
                        word_buf.seek(0)
                        send_tele(ID_USER_FULL, word_buf, f"Form_{name}.docx", f"✅ Engineer Log: {name}")
                        
                        excel_buf = sync_excel_master(excel_tmpl, edited_df)
                        send_tele(ID_USER_FULL, excel_buf, f"Master_Aiken_Update_{name}.xlsx", f"📊 Master Aiken Kumulatif")
                
                st.success(f"Execution Successful. {len(new_entries)} manual entries processed.")
                st.session_state.engine_df = edited_df.copy() # Reset original state

        except Exception as e:
            st.error(f"Execution Failed: {e}")

else:
    st.info("Run 'FETCH DATABASE' to start engine.")

st.markdown("---")
st.caption("Engineer Mode: Direct I/O | No UI Bloat | Batch Optimized")
