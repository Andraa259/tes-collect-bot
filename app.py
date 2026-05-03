import streamlit as st
import pandas as pd
from docx import Document
import gspread
from google.oauth2.service_account import Credentials
import io
import requests
import time

# --- 1. CONFIG ---
st.set_page_config(page_title="Engineer Master Hybrid v7.10", layout="wide")

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
    col_c = ws.col_values(3) 
    for r in range(4, 34): 
        if r > len(col_c) or not col_c[r-1]:
            return r
    return 34

def generate_word_final(name, job, skj, srel, skes, template_path):
    try:
        doc = Document(template_path)
        for p in doc.paragraphs:
            if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
            if "Pekerjaan\t:" in p.text: p.text = f"Pekerjaan\t: {job}"
        
        if doc.tables:
            table = doc.tables[0]
            for i in range(36):
                if i + 1 < len(table.rows):
                    row = table.rows[i + 1]
                    # Paksa jadi string untuk Word
                    row.cells[3].text = str(skj[i])
                    row.cells[4].text = str(srel[i])
                    row.cells[5].text = str(skes[i])
        
        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        return buf
    except Exception as e:
        st.error(f"Word Error: {e}"); return None

def send_tele(chat_id, file_buf, fname, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    file_buf.seek(0)
    requests.post(url, data={'chat_id': chat_id, 'caption': caption}, files={'document': (fname, file_buf)})

# --- 3. UI ---

st.title("🖥️ MASTER HYBRID v7.10")

if 'batch_data' not in st.session_state:
    st.session_state.batch_data = None

up_file = st.file_uploader("Upload Excel Aiken", type=["xlsx"])

if up_file:
    if st.button("⚙️ PROCESS & ALIGN DATA"):
        try:
            xl = pd.ExcelFile(up_file)
            temp_db = {}
            # Ambil Pekerjaan hanya dari KEJELASAN
            for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
                raw_xl = pd.read_excel(up_file, sheet_name=s_name, header=None, skiprows=3)
                raw_xl = raw_xl.reindex(columns=range(40), fill_value=0)
                
                rows = []
                for i in range(len(raw_xl)):
                    name = raw_xl.iloc[i, 1]
                    # Pekerjaan diambil di index 2 (Kolom C)
                    job = raw_xl.iloc[i, 2] if s_name == "KEJELASAN" else ""
                    
                    if pd.isna(name) or str(name).strip() == "": continue
                    
                    # Casting ke tipe data Python standar (Bukan Numpy)
                    row_dict = {
                        "Nama": str(name),
                        "Pekerjaan": str(job) if not pd.isna(job) else ""
                    }
                    
                    for a_idx in range(36):
                        val = raw_xl.iloc[i, 3 + a_idx]
                        # FIX: Paksa jadi int standar Python
                        try: row_dict[f"A{a_idx+1}"] = int(val) if not pd.isna(val) else 0
                        except: row_dict[f"A{a_idx+1}"] = 0
                    rows.append(row_dict)
                temp_db[s_name] = pd.DataFrame(rows)
            
            st.session_state.batch_data = temp_db
            st.success("Data Aligned! Pekerjaan diambil dari sheet KEJELASAN saja.")
            st.rerun()
        except Exception as e:
            st.error(f"Parsing Error: {e}")

if st.session_state.batch_data:
    tabs = st.tabs(["📊 KJ", "📈 REL", "📉 KES"])
    sheets = ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]
    for i, tab in enumerate(tabs):
        with tab:
            # Sembunyikan kolom Pekerjaan di REL dan KES biar ga bingung
            cols_to_show = st.session_state.batch_data[sheets[i]].columns
            if sheets[i] != "KEJELASAN":
                cols_to_show = [c for c in cols_to_show if c != "Pekerjaan"]
            
            st.session_state.batch_data[sheets[i]] = st.data_editor(
                st.session_state.batch_data[sheets[i]][cols_to_show],
                num_rows="dynamic",
                key=f"ed_{sheets[i]}"
            )

    if st.button("🚀 EXECUTE SYNC"):
        try:
            with st.spinner("Processing..."):
                client = get_gsheet_client()
                ss = client.open_by_url(GSHEET_URL)
                w_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
                
                db_kj = st.session_state.batch_data["KEJELASAN"]
                db_rel = st.session_state.batch_data["RELEVANSI"]
                db_kes = st.session_state.batch_data["KESESUAIAN"]

                for idx in range(len(db_kj)):
                    name = str(db_kj.iloc[idx]["Nama"])
                    # Ambil pekerjaan master dari KJ
                    job = str(db_kj.iloc[idx]["Pekerjaan"])
                    
                    # 1. Update GSheets
                    for s_name in sheets:
                        ws = ss.worksheet(s_name)
                        target_row = find_first_empty_row(ws)
                        if target_row <= 33:
                            # Ambil skor (kolom index 2 dst di dataframe)
                            current_df = st.session_state.batch_data[s_name]
                            # Filter kolom yang depannya 'A'
                            skor_cols = [c for c in current_df.columns if c.startswith('A')]
                            skor = current_df.iloc[idx][skor_cols].tolist()
                            
                            # Konversi list skor ke int standar Python lagi sebelum update
                            skor = [int(s) for s in skor]
                            
                            ws.update_cell(target_row, 2, name)
                            cells = ws.range(target_row, 3, target_row, 3 + len(skor) - 1)
                            for s_idx, val in enumerate(skor): cells[s_idx].value = val
                            ws.update_cells(cells)

                    # 2. Telegram Routing
                    skj_list = [int(s) for s in db_kj.iloc[idx][[c for c in db_kj.columns if c.startswith('A')]].tolist()]
                    srel_list = [int(s) for s in db_rel.iloc[idx][[c for c in db_rel.columns if c.startswith('A')]].tolist()]
                    skes_list = [int(s) for s in db_kes.iloc[idx][[c for c in db_kes.columns if c.startswith('A')]].tolist()]
                    
                    word_file = generate_word_final(name, job, skj_list, srel_list, skes_list, w_tmpl)
                    
                    if word_file:
                        # Kirim Word ke dua-duanya (karena ini dokumen utama)
                        send_tele(ID_USER_WORD, word_file, f"Form_{name}.docx", f"✅ {name}")
                        word_file.seek(0)
                        send_tele(ID_USER_FULL, word_file, f"Form_{name}.docx", f"✅ Full Log: {name}")

                # --- 3. KIRIM EXCEL KUMULATIF (HANYA KE ID_USER_FULL) ---
                # Logika ini bisa lo tambahin kalau lo punya file Excel yang mau dikirim di akhir
                st.success("Batching Berhasil!")
                st.session_state.batch_data = None
        except Exception as e:
            st.error(f"Pipeline Error: {e}")
