import streamlit as st
import pandas as pd
from docx import Document
import openpyxl
import gspread
from google.oauth2.service_account import Credentials
import io
import requests
import time

# --- 1. CONFIG ---
st.set_page_config(page_title="Engineer Injector v7.4", layout="wide")

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

# --- 3. CSS PROGRAMMER ---
st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #00ff41; }
    .stButton>button { border: 1px solid #00ff41 !important; color: #00ff41 !important; background: transparent !important; width: 100%; height: 3em; }
    .stTabs [data-baseweb="tab"] { color: #ffffff !important; }
    header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# --- 4. CORE FUNCTIONS ---

def get_gsheet_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(GCP_JSON, scopes=scope)
    return gspread.authorize(creds)

def fetch_all_sheets():
    """Tarik data dan BUANG semua baris hantu agar data baru nempel di atas"""
    try:
        client = get_gsheet_client()
        ss = client.open_by_url(GSHEET_URL)
        results = {}
        cols_gs = ["Nama"] + [f"A{i+1}" for i in range(36)]
        for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
            ws = ss.worksheet(s_name)
            # Ambil range B4:AL33
            raw_data = ws.get("B4:AL33")
            df = pd.DataFrame(raw_data, columns=cols_gs)
            
            # --- STRICT FILTERING ---
            # Hapus baris yang Namanya kosong/None/hanya spasi
            df = df.dropna(subset=["Nama"])
            df = df[df["Nama"].astype(str).str.strip() != ""]
            
            # Tambahkan kolom Pekerjaan (Virtual untuk Word)
            df.insert(1, "Pekerjaan", "") 
            results[s_name] = df.reset_index(drop=True)
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

# --- 5. APP ENGINE ---

if 'db' not in st.session_state:
    st.session_state.db = None

st.title("🖥️ BATCH INJECTOR v7.4 (FIX GHOST ROWS)")

# Tombol Fetch Utama
if st.session_state.db is None:
    if st.button("📥 INITIALIZE DATABASE (FETCH FROM GSHEETS)"):
        with st.spinner("Fetching clean data..."):
            st.session_state.db = fetch_all_sheets()
            if st.session_state.db:
                st.success("Connected! Data hantu baris 30+ sudah dibersihkan.")
                st.rerun()
    st.stop()

# --- MAIN INTERFACE ---
with st.sidebar:
    st.header("⚙️ Control")
    if st.button("🔄 RE-FETCH DATABASE"):
        st.session_state.db = fetch_all_sheets()
        st.rerun()

# 1. Uploader (Handle Excel lo "TESTCVI Aiken Zuyy Test.xlsx")
st.subheader("📂 Step 1: Upload Excel Baru")
up_file = st.file_uploader("Upload Master Excel", type=["xlsx"])

if up_file:
    if st.button("⚙️ MERGE UPLOAD"):
        try:
            xl = pd.ExcelFile(up_file)
            for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
                if s_name in xl.sheet_names:
                    # Sesuai file lo: Skip 3 baris header, Nama di kolom B (indeks 1)
                    # Karena di code v7 lo mau Pekerjaan, tapi di Excel lo gak ada, 
                    # kita sesuaikan: Nama (Col 1), Pekerjaan (Empty), Skor (Col 2-37)
                    raw_xl = pd.read_excel(up_file, sheet_name=s_name, header=None, skiprows=3)
                    
                    # Ambil Nama (Kolom indeks 1) dan Skor (Kolom indeks 2 s/d 37)
                    new_df = pd.DataFrame()
                    new_df["Nama"] = raw_xl.iloc[:, 1]
                    new_df["Pekerjaan"] = "" # Kosong karena di Excel lo gak ada
                    for i in range(36):
                        new_df[f"A{i+1}"] = raw_xl.iloc[:, 2 + i]
                    
                    # Bersihkan baris kosong dari hasil upload
                    new_df = new_df.dropna(subset=["Nama"])
                    new_df = new_df[new_df["Nama"].astype(str).str.strip() != ""]
                    
                    # Merge tepat di bawah data terakhir yang berisi nama
                    st.session_state.db[s_name] = pd.concat([st.session_state.db[s_name], new_df], ignore_index=True)
            
            st.success("Berhasil di-Merge! Sekarang data nempel tepat di baris atas.")
            st.rerun()
        except Exception as e:
            st.error(f"Excel Merge Error: {e}")

# 2. Review Grid
st.subheader("📝 Step 2: Review Data Grid")
tabs = st.tabs(["📊 KEJELASAN", "📈 RELEVANSI", "📉 KESESUAIAN"])
sheets = ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]

for i, tab in enumerate(tabs):
    with tab:
        st.session_state.db[sheets[i]] = st.data_editor(
            st.session_state.db[sheets[i]], 
            num_rows="dynamic", 
            key=f"editor_{sheets[i]}",
            use_container_width=True
        )

# 3. Execution
st.write("---")
if st.button("🚀 EXECUTE PIPELINE"):
    try:
        w_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
        with st.spinner("Syncing GSheets & Sending Word Documents..."):
            client = get_gsheet_client()
            ss = client.open_by_url(GSHEET_URL)
            
            # --- SYNC GSHEETS (OVERWRITE DARI B4) ---
            for s_name in sheets:
                ws = ss.worksheet(s_name)
                # Bersihkan dulu range B4:AL33 di GSheets biar gak sisa sampah
                empty_block = [["" for _ in range(37)] for _ in range(30)]
                ws.update("B4:AL33", empty_block)
                
                # Update dengan data baru
                df_to_gs = st.session_state.db[s_name].drop(columns=["Pekerjaan"])
                ws.update(f"B4:AL{3+len(df_to_gs)}", df_to_gs.fillna("").values.tolist())

            # --- KIRIM TELE ---
            db_kj = st.session_state.db["KEJELASAN"]
            db_rel = st.session_state.db["RELEVANSI"]
            db_kes = st.session_state.db["KESESUAIAN"]
            
            for idx in range(len(db_kj)):
                row = db_kj.iloc[idx]
                name, job = str(row["Nama"]), str(row["Pekerjaan"])
                if not name or name == "nan": continue
                
                skj = db_kj.iloc[idx, 2:].tolist()
                srel = db_rel.iloc[idx, 2:].tolist()
                skes = db_kes.iloc[idx, 2:].tolist()
                
                w_buf = generate_word_final(name, job, skj, srel, skes, w_tmpl)
                send_tele(ID_USER_WORD, w_buf, f"Form_{name}.docx", f"✅ Word: {name}")
                w_buf.seek(0)
                send_tele(ID_USER_FULL, w_buf, f"Form_{name}.docx", f"✅ Log: {name}")
            
            st.success("Pipeline Selesai!")
    except Exception as e:
        st.error(f"Pipeline Crash: {e}")
