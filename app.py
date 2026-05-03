import streamlit as st
import pandas as pd
from docx import Document
import openpyxl
import gspread
from google.oauth2.service_account import Credentials
import io
import requests
import time

# --- CONFIG & CREDENTIALS ---
TOKEN = st.secrets["TOKEN"]
ID_USER_WORD = st.secrets["CHAT_ID_1"]  # User 1: Word Only
ID_USER_FULL = st.secrets["CHAT_ID_2"]  # User 2: Word & Excel
GSHEET_URL = st.secrets["GSHEET_URL"]

# --- INITIAL STATE ---
if 'original_df' not in st.session_state:
    st.session_state.original_df = pd.DataFrame()

# --- CORE FUNCTIONS ---
def get_gsheet_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(creds)

def fetch_data():
    client = get_gsheet_client()
    ss = client.open_by_url(GSHEET_URL)
    ws = ss.worksheet("KEJELASAN")
    raw_data = ws.get("B4:AL33")
    cols = ["Nama"] + [f"Aitem_{i+1}" for i in range(36)]
    return pd.DataFrame(raw_data, columns=cols)

def send_to_tele(chat_id, file_stream, name, file_type="docx"):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    ext = ".docx" if file_type == "docx" else ".xlsx"
    fname = f"Form_Manual_{name}{ext}" if file_type == "docx" else f"CVI_Update_{name}{ext}"
    files = {'document': (fname, file_stream)}
    payload = {'chat_id': chat_id, 'caption': f"✅ Dokumen {file_type.upper()} - {name}"}
    requests.post(url, data=payload, files=files)

def generate_word(name, scores, template_path):
    doc = Document(template_path)
    for p in doc.paragraphs:
        if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
    
    if doc.tables:
        table = doc.tables[0]
        for i, score in enumerate(scores):
            if i + 1 < len(table.rows):
                row = table.rows[i + 1]
                for col in [3, 4, 5]: row.cells[col].text = str(score)
    
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

def process_excel_backup(template_path, current_df):
    """Menyalin seluruh data ke Excel Master untuk User 2"""
    wb = openpyxl.load_workbook(template_path)
    for sheet in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
        if sheet in wb.sheetnames:
            ws = wb[sheet]
            for r_idx, row in current_df.iterrows():
                target_r = 4 + r_idx
                if target_r > 33: break
                ws.cell(row=target_r, column=2, value=row["Nama"])
                for c_idx, val in enumerate(row[1:]):
                    try: ws.cell(row=target_r, column=3 + c_idx, value=int(val))
                    except: ws.cell(row=target_r, column=3 + c_idx, value=0)
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# --- UI INTERFACE ---
st.title("🚀 Hybrid Engine: Dual-Target Tele")
st.caption("Flow: Fetch -> Batch Input -> Sync GSheets -> Dual Tele Delivery")

with st.sidebar:
    st.header("⚙️ Resources")
    w_tmpl = st.file_uploader("Word Template", type="docx")
    e_tmpl = st.file_uploader("Excel Aiken Master", type="xlsx")
    if st.button("📥 1. Fetch Current GSheets"):
        st.session_state.original_df = fetch_data()
        st.success("Data Sinkron!")

if not st.session_state.original_df.empty:
    st.subheader("📝 Batch Data Input (Manual/Paste)")
    edited_df = st.data_editor(
        st.session_state.original_df,
        num_rows="dynamic",
        use_container_width=True,
        key="hybrid_editor"
    )

    if st.button("⚡ 2. OK: SYNC & SEND TO DUAL TELE"):
        if not w_tmpl or not e_tmpl:
            st.error("Upload Template Word & Excel dulu!")
        else:
            # Deteksi Baris Baru
            orig_len = len(st.session_state.original_df)
            new_data = edited_df.iloc[orig_len:]

            if new_data.empty:
                st.warning("Tidak ada data baru untuk di-generate.")
            else:
                with st.spinner("Processing Dual Delivery..."):
                    # 1. BACKUP SEMUA KE GSHEETS
                    client = get_gsheet_client()
                    ss = client.open_by_url(GSHEET_URL)
                    for s in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
                        ws = ss.worksheet(s)
                        ws.update(f"B4:AL{3+len(edited_df)}", edited_df.fillna("").values.tolist())

                    # 2. GENERATE & SEND (NEW DATA ONLY)
                    for i, (idx, row) in enumerate(new_data.iterrows()):
                        p_name = str(row["Nama"])
                        p_scores = row[1:].tolist()
                        
                        # Generate Word
                        word_buf = generate_word(p_name, p_scores, w_tmpl)
                        
                        # USER 1: Kirim Word Saja
                        send_to_tele(ID_USER_WORD, word_buf, p_name, "docx")
                        
                        # USER 2: Kirim Word & Master Excel Terupdate
                        word_buf.seek(0) # Reset stream
                        send_to_tele(ID_USER_FULL, word_buf, p_name, "docx")
                        
                        excel_buf = process_excel_backup(e_tmpl, edited_df)
                        send_to_tele(ID_USER_FULL, excel_buf, p_name, "xlsx")
                    
                    st.success(f"Berhasil! Sync GSheets OK & {len(new_data)} data baru dikirim ke 2 User Tele.")
                    st.session_state.original_df = edited_df.copy()
