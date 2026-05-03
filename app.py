import streamlit as st
import pandas as pd
from docx import Document
import openpyxl
import gspread
from google.oauth2.service_account import Credentials
import io
import requests
import time

# --- 1. CONFIG (MUST BE FIRST) ---
st.set_page_config(page_title="Engineer Injector v7.3", layout="wide")

# --- 2. CREDENTIALS ---
try:
    TOKEN = st.secrets["TOKEN"]
    ID_USER_WORD = st.secrets.get("CHAT_ID_1")
    ID_USER_FULL = st.secrets.get("CHAT_ID_2")
    GSHEET_URL = st.secrets["GSHEET_URL"]
    GCP_JSON = st.secrets["gcp_service_account"]
except Exception as e:
    st.error(f"Secret Error: {e}")
    st.stop()

# --- 3. CSS (CLEANED) ---
st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #00ff41; }
    .stButton>button { border: 1px solid #00ff41 !important; color: #00ff41 !important; background: transparent !important; width: 100%; height: 3em; }
    /* Memastikan teks di tab kelihatan di mobile */
    .stTabs [data-baseweb="tab"] { color: #ffffff !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 4. CORE FUNCTIONS ---

def get_gsheet_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(GCP_JSON, scopes=scope)
    return gspread.authorize(creds)

def fetch_all_sheets():
    try:
        client = get_gsheet_client()
        ss = client.open_by_url(GSHEET_URL)
        results = {}
        cols_gs = ["Nama"] + [f"A{i+1}" for i in range(36)]
        for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
            ws = ss.worksheet(s_name)
            data = ws.get("B4:AL33")
            df = pd.DataFrame(data, columns=cols_gs)
            df = df[df["Nama"].astype(str).str.strip() != ""]
            df.insert(1, "Pekerjaan", "") 
            results[s_name] = df
        return results
    except Exception as e:
        st.error(f"Fetch Error: {e}")
        return None

def send_tele(chat_id, file_buf, fname, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    file_buf.seek(0)
    return requests.post(url, data={'chat_id': chat_id, 'caption': caption}, files={'document': (fname, file_buf)})

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
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- 5. MAIN LOGIC ---

# Initialize State
if 'db' not in st.session_state:
    st.session_state.db = None

# Tombol Emergency di Main Page (Kalau Sidebar ngumpet di Android)
st.title("🖥️ BATCH INJECTOR v7.3")

if st.session_state.db is None:
    st.warning("Data belum ditarik dari GSheets.")
    if st.button("📥 FETCH & INITIALIZE DATABASE"):
        with st.spinner("Fetching..."):
            st.session_state.db = fetch_all_sheets()
            if st.session_state.db:
                st.success("Data loaded!")
                st.rerun()
    st.stop()

# --- SIDEBAR (Tetap Ada Buat Control) ---
with st.sidebar:
    st.header("⚙️ Control Panel")
    if st.button("🔄 RE-FETCH GSHEETS"):
        st.session_state.db = fetch_all_sheets()
        st.rerun()

# --- MAIN CONTENT ---

# 1. Uploader
st.subheader("📂 Step 1: Upload Excel Baru")
up_file = st.file_uploader("Format Sheet: KEJELASAN, RELEVANSI, KESESUAIAN", type=["xlsx"])

if up_file:
    if st.button("⚙️ MERGE DATA"):
        try:
            xl = pd.ExcelFile(up_file)
            for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
                if s_name in xl.sheet_names:
                    new_data = xl.parse(s_name, header=None)
                    new_data.columns = ["Nama", "Pekerjaan"] + [f"A{i+1}" for i in range(36)]
                    st.session_state.db[s_name] = pd.concat([st.session_state.db[s_name], new_data], ignore_index=True)
            st.success("Merged!")
            st.rerun()
        except Exception as e:
            st.error(f"Excel Error: {e}")

# 2. Grid Editor
st.subheader("📝 Step 2: Review Data")
tabs = st.tabs(["📊 KJ", "📈 REL", "📉 KES"])
sheets = ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]

for i, tab in enumerate(tabs):
    with tab:
        # Gunakan try-except untuk handle versi Streamlit lama
        try:
            st.session_state.db[sheets[i]] = st.data_editor(
                st.session_state.db[sheets[i]], 
                num_rows="dynamic", 
                key=f"editor_{sheets[i]}",
                use_container_width=True
            )
        except:
            st.session_state.db[sheets[i]] = st.data_editor(
                st.session_state.db[sheets[i]], 
                num_rows="dynamic", 
                key=f"editor_alt_{sheets[i]}"
            )

# 3. Execution
st.write("---")
if st.button("🚀 EXECUTE PIPELINE"):
    try:
        w_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
        db_kj = st.session_state.db["KEJELASAN"]
        db_rel = st.session_state.db["RELEVANSI"]
        db_kes = st.session_state.db["KESESUAIAN"]
        
        # Logic deteksi baris baru bisa lo tambahin di sini atau hajar semua
        # Kita hajar yang Namanya valid aja
        valid_rows = db_kj[db_kj["Nama"].astype(str).str.strip() != ""]
        
        with st.spinner(f"Processing {len(valid_rows)} records..."):
            client = get_gsheet_client()
            ss = client.open_by_url(GSHEET_URL)
            
            # Update GSheets
            for s_name in sheets:
                ws = ss.worksheet(s_name)
                df_to_gs = st.session_state.db[s_name].drop(columns=["Pekerjaan"])
                ws.update(f"B4:AL{3+len(df_to_gs)}", df_to_gs.fillna("").values.tolist())

            # Kirim Tele
            for idx in range(len(valid_rows)):
                row = valid_rows.iloc[idx]
                name, job = str(row["Nama"]), str(row["Pekerjaan"])
                
                # Cari skor di 3 sheet berdasarkan index
                skj = db_kj.iloc[idx, 2:].tolist()
                srel = db_rel.iloc[idx, 2:].tolist()
                skes = db_kes.iloc[idx, 2:].tolist()
                
                w_buf = generate_word_final(name, job, skj, srel, skes, w_tmpl)
                send_tele(ID_USER_WORD, w_buf, f"Form_{name}.docx", f"✅ Word: {name}")
                w_buf.seek(0)
                send_tele(ID_USER_FULL, w_buf, f"Form_{name}.docx", f"✅ Log: {name}")
            
            st.success("Semua Data Berhasil Diproses!")
    except Exception as e:
        st.error(f"Pipeline Crash: {e}")
