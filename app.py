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
st.set_page_config(page_title="Engineer Injector v7.5", layout="wide")

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

# --- 3. CSS ---
st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #00ff41; }
    .stButton>button { border: 1px solid #00ff41 !important; color: #00ff41 !important; background: transparent !important; width: 100%; height: 3em; }
    header {visibility: hidden;}
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
            raw_data = ws.get("B4:AL33")
            df = pd.DataFrame(raw_data, columns=cols_gs)
            df = df.dropna(subset=["Nama"])
            df = df[df["Nama"].astype(str).str.strip() != ""]
            df.insert(1, "Pekerjaan", "") 
            results[s_name] = df.reset_index(drop=True)
        return results
    except Exception as e:
        st.error(f"Fetch GSheets Error: {e}")
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
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 5. APP ENGINE ---

if 'db' not in st.session_state:
    st.session_state.db = None

st.title("🖥️ BATCH INJECTOR v7.5 (ROBUST INDEXER)")

if st.session_state.db is None:
    if st.button("📥 INITIALIZE DATABASE"):
        st.session_state.db = fetch_all_sheets()
        if st.session_state.db: st.rerun()
    st.stop()

with st.sidebar:
    if st.button("🔄 RE-FETCH DATABASE"):
        st.session_state.db = fetch_all_sheets()
        st.rerun()

# 1. Uploader
st.subheader("📂 Step 1: Upload Excel")
up_file = st.file_uploader("Upload Master Excel", type=["xlsx"])

if up_file:
    if st.button("⚙️ MERGE UPLOAD"):
        try:
            xl = pd.ExcelFile(up_file)
            for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
                if s_name in xl.sheet_names:
                    # Baca data mulai baris 4 (Testing)
                    raw_xl = pd.read_excel(up_file, sheet_name=s_name, header=None, skiprows=3)
                    
                    # Logic Robust: Ambil kolom Nama (Indeks 1)
                    # Ambil semua kolom skor mulai dari Indeks 2
                    names = raw_xl.iloc[:, 1].values
                    scores_part = raw_xl.iloc[:, 2:]
                    
                    new_rows = []
                    for i, name in enumerate(names):
                        if pd.isna(name) or str(name).strip() == "":
                            continue
                        
                        # Buat dictionary baris baru
                        row_dict = {"Nama": name, "Pekerjaan": ""}
                        # Ambil skor yang tersedia, kalau kurang dari 36 isi 0
                        for a_idx in range(36):
                            try:
                                # Ambil nilai kolom ke i+2
                                val = scores_part.iloc[i, a_idx]
                                row_dict[f"A{a_idx+1}"] = val if not pd.isna(val) else 0
                            except IndexError:
                                # Jika kolom tidak ada di Excel, isi 0
                                row_dict[f"A{a_idx+1}"] = 0
                        new_rows.append(row_dict)
                    
                    new_df = pd.DataFrame(new_rows)
                    st.session_state.db[s_name] = pd.concat([st.session_state.db[s_name], new_df], ignore_index=True)
            
            st.success("Merge Sukses! Kolom yang hilang otomatis diisi 0.")
            st.rerun()
        except Exception as e:
            st.error(f"Logic Error: {e}")

# 2. Grid Editor
st.subheader("📝 Step 2: Review Data")
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
if st.button("🚀 EXECUTE PIPELINE"):
    try:
        w_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
        with st.spinner("Processing..."):
            client = get_gsheet_client(); ss = client.open_by_url(GSHEET_URL)
            for s_name in sheets:
                ws = ss.worksheet(s_name)
                # Reset range biar bersih
                empty = [["" for _ in range(37)] for _ in range(30)]
                ws.update("B4:AL33", empty)
                df_to_gs = st.session_state.db[s_name].drop(columns=["Pekerjaan"])
                ws.update(f"B4:AL{3+len(df_to_gs)}", df_to_gs.fillna("").values.tolist())

            db_kj, db_rel, db_kes = st.session_state.db["KEJELASAN"], st.session_state.db["RELEVANSI"], st.session_state.db["KESESUAIAN"]
            for idx in range(len(db_kj)):
                row = db_kj.iloc[idx]
                if pd.isna(row["Nama"]) or str(row["Nama"]).strip() == "": continue
                w_buf = generate_word_final(row["Nama"], row["Pekerjaan"], db_kj.iloc[idx, 2:].tolist(), db_rel.iloc[idx, 2:].tolist(), db_kes.iloc[idx, 2:].tolist(), w_tmpl)
                send_tele(ID_USER_WORD, w_buf, f"Form_{row['Nama']}.docx", f"✅ {row['Nama']}")
                w_buf.seek(0); send_tele(ID_USER_FULL, w_buf, f"Form_{row['Nama']}.docx", f"✅ Log: {row['Nama']}")
            st.success("Selesai!")
    except Exception as e:
        st.error(f"Pipeline Crash: {e}")
