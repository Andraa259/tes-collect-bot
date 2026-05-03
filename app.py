import streamlit as st
import pandas as pd
from docx import Document
import gspread
from google.oauth2.service_account import Credentials
import io
import requests
import openpyxl
import time

# --- 1. CONFIG & SECRETS ---
st.set_page_config(page_title="Engineer Master Hybrid v7.11", layout="wide")

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

def generate_word_final(name, job, master_scores, template_path):
    """
    Injeksi skor menggunakan Logika Pencocokan Teks (Text Matching)
    agar skor tidak salah letak meskipun template punya banyak header.
    """
    try:
        doc = Document(template_path)
        # 1. Injeksi Identitas
        for p in doc.paragraphs:
            if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
            if "Pekerjaan\t:" in p.text: p.text = f"Pekerjaan\t: {job}"
        
        # 2. Injeksi Skor via Match Text
        if doc.tables:
            table = doc.tables[0]
            for row in table.rows:
                # Normalisasi teks di kolom Aitem (Cell index 2)
                cell_text = "".join(row.cells[2].text.split()).lower()
                
                # Cari kecocokan dengan data aitem
                for item_text, scores in master_scores.items():
                    target_match = "".join(item_text.split()).lower()[:60]
                    if target_match in cell_text:
                        row.cells[3].text = str(scores['kj'])
                        row.cells[4].text = str(scores['rel'])
                        row.cells[5].text = str(scores['kes'])
        
        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        return buf
    except Exception as e:
        st.error(f"Word Engine Error: {e}"); return None

def proses_excel_cvi():
    """Fungsi Batching untuk menghasilkan Excel Kumulatif dari GSheets"""
    try:
        client = get_gsheet_client(); ss = client.open_by_url(GSHEET_URL)
        # Pastikan file template "CVI Aiken Zuyy.xlsx" ada di folder yang sama
        wb = openpyxl.load_workbook("CVI Aiken Zuyy.xlsx")
        for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
            ws_gs = ss.worksheet(s_name)
            all_vals = ws_gs.get_all_values()
            ws_xl = wb[s_name]
            # Batching data dari GSheets baris 4 dst ke Excel
            for idx, row_data in enumerate(all_vals[3:]):
                target_row = 4 + idx
                if target_row > 33 or len(row_data) < 2 or not row_data[1]: break
                ws_xl.cell(row=target_row, column=2, value=row_data[1]) # Nama
                for col_idx, val in enumerate(row_data[2:38]): # Skor 1-36
                    try: ws_xl.cell(row=target_row, column=3+col_idx, value=int(float(val)))
                    except: ws_xl.cell(row=target_row, column=3+col_idx, value=0)
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return buf
    except Exception as e:
        st.error(f"CVI Excel Error: {e}"); return None

def send_tele(chat_id, file_buf, fname, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    file_buf.seek(0)
    requests.post(url, data={'chat_id': chat_id, 'caption': caption}, files={'document': (fname, file_buf)})

# --- 3. UI ENGINE ---

st.title("🖥️ MASTER HYBRID v7.11 (SMART ALIGN)")

if 'batch_data' not in st.session_state:
    st.session_state.batch_data = None

up_file = st.file_uploader("Upload Excel Aiken", type=["xlsx"])

if up_file:
    if st.button("⚙️ PROCESS & ALIGN DATA"):
        try:
            xl = pd.ExcelFile(up_file)
            temp_db = {}
            for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
                raw_xl = pd.read_excel(up_file, sheet_name=s_name, header=None, skiprows=3)
                raw_xl = raw_xl.reindex(columns=range(40), fill_value=0)
                
                rows = []
                for i in range(len(raw_xl)):
                    name = raw_xl.iloc[i, 1]
                    # Pekerjaan hanya diambil dari sheet KEJELASAN
                    job = raw_xl.iloc[i, 2] if s_name == "KEJELASAN" else ""
                    if pd.isna(name) or str(name).strip() == "": continue
                    
                    row_dict = {"Nama": str(name), "Pekerjaan": str(job) if not pd.isna(job) else ""}
                    for a_idx in range(36):
                        val = raw_xl.iloc[i, 3 + a_idx]
                        # Casting to standard Python Int
                        try: row_dict[f"A{a_idx+1}"] = int(val) if not pd.isna(val) else 0
                        except: row_dict[f"A{a_idx+1}"] = 0
                    rows.append(row_dict)
                temp_db[s_name] = pd.DataFrame(rows)
            
            st.session_state.batch_data = temp_db
            st.success("Excel Aligned! Pekerjaan Master terkunci di sheet KEJELASAN.")
            st.rerun()
        except Exception as e:
            st.error(f"Parsing Error: {e}")

if st.session_state.batch_data:
    tabs = st.tabs(["📊 KJ", "📈 REL", "📉 KES"])
    sheets = ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]
    for i, tab in enumerate(tabs):
        with tab:
            df_display = st.session_state.batch_data[sheets[i]]
            # Sembunyikan kolom Pekerjaan di sheet selain KJ
            if sheets[i] != "KEJELASAN":
                df_display = df_display.drop(columns=["Pekerjaan"])
            st.session_state.batch_data[sheets[i]] = st.data_editor(df_display, num_rows="dynamic", key=f"ed_{sheets[i]}")

    if st.button("🚀 EXECUTE SMART SYNC"):
        try:
            with st.spinner("Syncing GSheets, Word & Aiken Kumulatif..."):
                client = get_gsheet_client(); ss = client.open_by_url(GSHEET_URL)
                w_tmpl = "Form Validasi Expert Judgement Ayinn Ver. 3.docx"
                
                db_kj = st.session_state.batch_data["KEJELASAN"]
                db_rel = st.session_state.batch_data["RELEVANSI"]
                db_kes = st.session_state.batch_data["KESESUAIAN"]

                for idx in range(len(db_kj)):
                    name = str(db_kj.iloc[idx]["Nama"])
                    job = str(db_kj.iloc[idx]["Pekerjaan"])
                    
                    # 1. Update GSheets (Smart Append)
                    for s_name in sheets:
                        ws = ss.worksheet(s_name)
                        target_row = find_first_empty_row(ws)
                        if target_row <= 33:
                            current_df = st.session_state.batch_data[s_name]
                            skor_cols = [c for c in current_df.columns if c.startswith('A')]
                            skor = [int(s) for s in current_df.iloc[idx][skor_cols].tolist()]
                            
                            ws.update_cell(target_row, 2, name)
                            cells = ws.range(target_row, 3, target_row, 3 + len(skor) - 1)
                            for s_idx, val in enumerate(skor): cells[s_idx].value = val
                            ws.update_cells(cells)

                    # 2. Word Injection (Menggunakan Map untuk Akurasi)
                    # Kita buat peta: "Teks Aitem" -> {kj, rel, kes}
                    aitem_texts = [c for c in db_kj.columns if c.startswith('A')]
                    # Catatan: Karena di batch kita cuma pake label A1-A36, lo butuh list teks aslinya
                    # Gue asumsikan lo pake teks dari variabel data_aspek atau sejenisnya.
                    # Tapi biar aman, kita pake urutan aitem sesuai list di Word.
                    
                    # Kita kumpulin skor per aitem
                    skj_list = db_kj.iloc[idx][[c for c in db_kj.columns if c.startswith('A')]].tolist()
                    srel_list = db_rel.iloc[idx][[c for c in db_rel.columns if c.startswith('A')]].tolist()
                    skes_list = db_kes.iloc[idx][[c for c in db_kes.columns if c.startswith('A')]].tolist()
                    
                    # Peta Skor sederhana untuk Word
                    # Jika teks di Word tidak ketemu, ini fallback-nya
                    master_scores = {}
                    # Disini gue pake trik: kirim list skor saja, generate_word_final bakal tetep pake 
                    # i+1 tapi gue tambahin proteksi pencocokan baris "1." dst.
                    
                    word_file = generate_word_final(name, job, {}, w_tmpl) # Fallback placeholder
                    
                    # RE-FIX Word Logic buat Batch: Kita pake baris tabel yang depannya ada Angka.
                    doc_final = Document(w_tmpl)
                    for p in doc_final.paragraphs:
                        if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {name}"
                        if "Pekerjaan\t:" in p.text: p.text = f"Pekerjaan\t: {job}"
                    
                    if doc_final.tables:
                        tbl = doc_final.tables[0]
                        aitem_idx = 0
                        for row in tbl.rows:
                            # Cek apakah kolom pertama (No) berisi angka urut (1., 2., dst)
                            first_cell = row.cells[0].text.strip()
                            if first_cell.endswith(".") or first_cell.isdigit():
                                if aitem_idx < 36:
                                    row.cells[3].text = str(skj_list[aitem_idx])
                                    row.cells[4].text = str(srel_list[aitem_idx])
                                    row.cells[5].text = str(skes_list[aitem_idx])
                                    aitem_idx += 1
                    
                    word_buf = io.BytesIO(); doc_final.save(word_buf); word_buf.seek(0)
                    
                    if word_buf:
                        send_tele(ID_USER_WORD, word_buf, f"Form Validasi Expert Judgement Forgiveness_{name}.docx", f"✅ Word: {name}")
                        word_buf.seek(0)
                        send_tele(ID_USER_FULL, word_buf, f"Form Validasi Expert Judgement Forgiveness_{name}.docx", f"✅ Full Log: {name}")

                # 3. KIRIM EXCEL KUMULATIF KE ID_USER_FULL
                cvi_buf = proses_excel_cvi()
                if cvi_buf:
                    send_tele(ID_USER_FULL, cvi_buf, f"CVI_Aiken_Update_{int(time.time())}.xlsx", "📊 Rekap Aiken Kumulatif Terbaru")
                
                st.success("Pipeline Selesai: GSheets Terupdate, Word Terkirim, & Rekap Aiken Terkirim ke Admin!")
                st.session_state.batch_data = None
        except Exception as e:
            st.error(f"Pipeline Error: {e}")
