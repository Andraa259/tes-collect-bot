import streamlit as st
import pandas as pd
from docx import Document
import gspread
from google.oauth2.service_account import Credentials
import io
import requests
import openpyxl
import time

# --- 1. CONFIG ---
st.set_page_config(page_title="Engineer Master Hybrid v7.14", layout="wide")

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

def proses_excel_cvi():
    """Mengambil data TERBARU dari GSheets dan merakitnya ke file Excel Kumulatif"""
    try:
        client = get_gsheet_client()
        ss = client.open_by_url(GSHEET_URL)
        # Pastikan file 'CVI Aiken Zuyy.xlsx' ada di folder project GitHub lo
        wb = openpyxl.load_workbook("CVI Aiken Zuyy.xlsx")
        
        for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
            ws_gs = ss.worksheet(s_name)
            all_vals = ws_gs.get_all_values() # Ambil semua data GSheets
            ws_xl = wb[s_name]
            
            # Looping mulai baris 4 GSheets (index 3)
            for idx, row_data in enumerate(all_vals[3:]):
                target_row = 4 + idx
                if target_row > 33 or len(row_data) < 2 or not row_data[1]: 
                    break
                
                # Isi Nama ke Kolom B
                ws_xl.cell(row=target_row, column=2, value=row_data[1])
                
                # Isi Skor A1-A36 ke Kolom C-AL
                for col_idx, val in enumerate(row_data[2:38]):
                    try:
                        # Convert ke angka biar Excel bisa ngitung formulanya
                        ws_xl.cell(row=target_row, column=3 + col_idx, value=int(float(val)))
                    except:
                        ws_xl.cell(row=target_row, column=3 + col_idx, value=0)
                        
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf
    except Exception as e:
        st.error(f"Gagal generate Excel Kumulatif: {e}")
        return None

def send_tele(chat_id, file_buf, fname, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    file_buf.seek(0)
    return requests.post(url, data={'chat_id': chat_id, 'caption': caption}, files={'document': (fname, file_buf)})

# --- 3. UI ENGINE ---

st.title("🖥️ MASTER HYBRID v7.14 (CUMULATIVE SYNC)")

if 'batch_data' not in st.session_state:
    st.session_state.batch_data = None

up_file = st.file_uploader("Upload Excel Aiken", type=["xlsx"])

if up_file:
    if st.button("⚙️ PROCESS DATA"):
        try:
            xl = pd.ExcelFile(up_file)
            temp_db = {}
            for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
                raw_xl = pd.read_excel(up_file, sheet_name=s_name, header=None, skiprows=3)
                raw_xl = raw_xl.reindex(columns=range(40), fill_value=0)
                
                rows = []
                for i in range(len(raw_xl)):
                    name = raw_xl.iloc[i, 1] 
                    if pd.isna(name) or str(name).strip() == "": continue
                    
                    if s_name == "KEJELASAN":
                        job = raw_xl.iloc[i, 2] 
                        start_skor = 3 
                    else:
                        job = "" 
                        start_skor = 2 
                    
                    row_dict = {"Nama": str(name), "Pekerjaan": str(job) if not pd.isna(job) else ""}
                    for a_idx in range(36):
                        val = raw_xl.iloc[i, start_skor + a_idx]
                        try: row_dict[f"A{a_idx+1}"] = int(val) if not pd.isna(val) else 0
                        except: row_dict[f"A{a_idx+1}"] = 0
                    rows.append(row_dict)
                temp_db[s_name] = pd.DataFrame(rows)
            
            st.session_state.batch_data = temp_db
            st.success("Excel Terbaca! Skor diatur otomatis sesuai struktur sheet.")
            st.rerun()
        except Exception as e:
            st.error(f"Error: {e}")

if st.session_state.batch_data:
    tabs = st.tabs(["📊 KJ", "📈 REL", "📉 KES"])
    sheets = ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]
    
    for i, tab in enumerate(tabs):
        with tab:
            config = {}
            if sheets[i] != "KEJELASAN": config["Pekerjaan"] = None
            
            st.session_state.batch_data[sheets[i]] = st.data_editor(
                st.session_state.batch_data[sheets[i]],
                num_rows="dynamic",
                key=f"ed_{sheets[i]}",
                column_config=config
            )

    if st.button("🚀 EXECUTE SYNC"):
        try:
            with st.spinner("Processing GSheets & Word..."):
                client = get_gsheet_client()
                ss = client.open_by_url(GSHEET_URL)
                w_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
                
                db_kj = st.session_state.batch_data["KEJELASAN"]
                db_rel = st.session_state.batch_data["RELEVANSI"]
                db_kes = st.session_state.batch_data["KESESUAIAN"]

                for idx in range(len(db_kj)):
                    name = str(db_kj.iloc[idx]["Nama"])
                    job = str(db_kj.iloc[idx]["Pekerjaan"]) 
                    
                    for s_name in sheets:
                        ws = ss.worksheet(s_name)
                        target_row = find_first_empty_row(ws)
                        if target_row <= 33:
                            curr_df = st.session_state.batch_data[s_name]
                            skor = [int(curr_df.iloc[idx][f"A{k+1}"]) for k in range(36)]
                            
                            ws.update_cell(target_row, 2, name)
                            cells = ws.range(target_row, 3, target_row, 3 + len(skor) - 1)
                            for s_idx, v in enumerate(skor): cells[s_idx].value = v
                            ws.update_cells(cells)

                    # Generate Word
                    doc = Document(w_tmpl)
                    for p in doc.paragraphs:
                        if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
                        if "Pekerjaan\t:" in p.text: p.text = f"Pekerjaan\t: {job}"
                    
                    if doc.tables:
                        table = doc.tables[0]
                        a_idx = 0
                        skj = [int(db_kj.iloc[idx][f"A{k+1}"]) for k in range(36)]
                        srel = [int(db_rel.iloc[idx][f"A{k+1}"]) for k in range(36)]
                        skes = [int(db_kes.iloc[idx][f"A{k+1}"]) for k in range(36)]
                        
                        for row in table.rows:
                            txt_no = row.cells[0].text.strip()
                            if (txt_no.endswith(".") or txt_no.isdigit()) and a_idx < 36:
                                row.cells[3].text, row.cells[4].text, row.cells[5].text = str(skj[a_idx]), str(srel[a_idx]), str(skes[a_idx])
                                a_idx += 1
                    
                    word_buf = io.BytesIO(); doc.save(word_buf); word_buf.seek(0)
                    send_tele(ID_USER_WORD, word_buf, f"Form Validasi Expert Judgement Forgiveness_{name}.docx", f"✅ Word: {name}")
                    word_buf.seek(0)
                    send_tele(ID_USER_FULL, word_buf, f"Form Validasi Expert Judgement Forgiveness_{name}.docx", f"✅ Full Log: {name}")

                # --- 🎯 BAGIAN PENGIRIMAN EXCEL KUMULATIF ---
                st.write("---")
                with st.spinner("Meresume GSheets ke Excel Kumulatif..."):
                    excel_kumulatif = proses_excel_cvi()
                    if excel_kumulatif:
                        send_tele(ID_USER_FULL, excel_kumulatif, f"Rekap_CVI_Aiken_Latest_{int(time.time())}.xlsx", "📊 REKAP KUMULATIF SELURUH PANELIS")
                        st.success("Excel Kumulatif terkirim ke Admin!")

                st.success("Selesai! GSheets terisi & Word terkirim.")
                st.session_state.batch_data = None
        except Exception as e:
            st.error(f"Pipeline Error: {e}")
