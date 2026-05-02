import streamlit as st
from docx import Document
import requests
import io
import openpyxl
import gspread
from google.oauth2.service_account import Credentials
from streamlit_scroll_to_top import scroll_to_here
import time

# --- CONFIG & SECRETS ---
TOKEN = st.secrets["TOKEN"]
ID_USER_WORD = st.secrets["CHAT_ID_1"]  # User 1: Word Only
ID_USER_FULL = st.secrets["CHAT_ID_2"]  # User 2: Word & Excel
GSHEET_URL = st.secrets["GSHEET_URL"]

# --- INITIALIZING SESSION STATE ---
if 'step' not in st.session_state: st.session_state.step = 0
if 'scroll_to_top' not in st.session_state: st.session_state.scroll_to_top = False
if 'master_data' not in st.session_state: st.session_state.master_data = {}
if 'p_nama' not in st.session_state: st.session_state.p_nama = ""
if 'p_kerja' not in st.session_state: st.session_state.p_kerja = ""
if 's_global' not in st.session_state: st.session_state.s_global = ""
if 'submitted' not in st.session_state: st.session_state.submitted = False
if 'confirm_send' not in st.session_state: st.session_state.confirm_send = False

# --- OPTIMASI SCROLL ---
if st.session_state.scroll_to_top:
    scroll_to_here(0, key=f'scroll_trigger_{st.session_state.step}')
    st.session_state.scroll_to_top = False

def move_step(num):
    st.session_state.scroll_to_top = True
    st.session_state.step = num

# --- FUNGSI MULTI-TARGET TELEGRAM ---
def kirim_telegram_multi(word_buf, excel_buf, nama):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    targets = [
        {"id": ID_USER_WORD, "files": [("docx", word_buf)]},
        {"id": ID_USER_FULL, "files": [("docx", word_buf), ("xlsx", excel_buf)]}
    ]
    for t in targets:
        for f_type, f_buf in t["files"]:
            if f_buf:
                f_buf.seek(0)
                fname = f"Form_{nama}.docx" if f_type == "docx" else f"CVI_Aiken_{nama}.xlsx"
                requests.post(url, data={'chat_id': t["id"], 'caption': f"✅ {f_type.upper()} Masuk: {nama}"}, files={'document': (fname, f_buf)})

# --- GSHEETS & EXCEL LOGIC (KUNCIAN C4-C33) ---
def get_gs_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(creds)

def simpan_ke_gsheets():
    try:
        ss = get_gs_client().open_by_url(GSHEET_URL)
        kategori = {"KEJELASAN": "kj", "RELEVANSI": "rel", "KESESUAIAN": "kes"}
        for s_name, k_data in kategori.items():
            ws = ss.worksheet(s_name)
            col_c = ws.col_values(3)
            tr = 4
            for r in range(4, 34):
                if r > len(col_c) or not col_c[r-1]:
                    tr = r; break
                tr = r + 1
            if tr <= 33:
                ws.update_cell(tr, 2, st.session_state.p_nama)
                skor = []
                for asp in ["Pemaafan Diri", "Pemaafan Orang Lain", "Pemaafan Situasi"]:
                    for _, items in data_aspek[asp]:
                        for txt in items: skor.append(st.session_state.master_data.get(txt, {k_data: 0})[k_data])
                cells = ws.range(tr, 3, tr, 3 + len(skor) - 1)
                for i, s in enumerate(skor): cells[i].value = s
                ws.update_cells(cells)
        return True
    except: return False

def proses_excel_cvi():
    try:
        ss = get_gs_client().open_by_url(GSHEET_URL)
        wb = openpyxl.load_workbook("CVI Aiken Zuyy.xlsx")
        for s_name in ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]:
            ws_gs, ws_xl = ss.worksheet(s_name), wb[s_name]
            all_v = ws_gs.get_all_values()
            for idx, r_data in enumerate(all_v[3:]):
                tr = 4 + idx
                if tr > 33 or not r_data[2]: break
                ws_xl.cell(row=tr, column=2, value=r_data[1])
                for ci, v in enumerate(r_data[2:38]):
                    try: ws_xl.cell(row=tr, column=3 + ci, value=int(float(v)))
                    except: ws_xl.cell(row=tr, column=3 + ci, value=0)
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return buf
    except: return None

# --- UI STYLING & RESPONSIVE ---
st.set_page_config(page_title="Expert Judgement", layout="centered")
st.markdown("""
    <style>
    .intro-card { background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); color: white; padding: 40px; border-radius: 20px; text-align: center; margin-bottom: 20px; }
    .indicator-header { background-color: #1E3A8A; color: white; padding: 12px; border-radius: 10px 10px 0 0; font-weight: bold; text-align: center; margin-top: 15px; }
    .white-card { background-color: rgba(255,255,255,0.9); border: 1px solid #E2E8F0; padding: 20px; border-radius: 0 0 10px 10px; margin-bottom: 20px; }
    .completed-card { border-left: 8px solid #10B981 !important; background-color: #F0FDF4 !important; }
    .stButton>button { border-radius: 12px; height: 50px; font-weight: bold; width: 100%; }
    @media (max-width: 640px) { .white-card { padding: 15px; } .intro-card { padding: 25px; } }
    </style>
    """, unsafe_allow_html=True)

data_aspek = {
    "Pemaafan Diri": [ ("Indikator 1", ["Item 1...", "Item 2..."]), ("Indikator 2", ["Item 3...", "Item 4..."]) ],
    "Pemaafan Orang Lain": [ ("Indikator 3", ["Item 5...", "Item 6..."]) ],
    "Pemaafan Situasi": [ ("Indikator 4", ["Item 7...", "Item 8..."]) ]
} # Selesaikan isi data_aspek sesuai instrumenmu

# --- LOGIKA PROGRESS BAR ---
total_scoring_steps = 3 # 3 aspek
if 2 <= st.session_state.step <= 4:
    progress = (st.session_state.step - 1) / (total_scoring_steps + 1)
    st.progress(progress)

# --- ALUR APLIKASI ---
if st.session_state.step == 0:
    st.markdown("<div class='intro-card'><h1>Expert Judgement Portal</h1><p>Psikometri & Validasi Instrumen</p></div>", unsafe_allow_html=True)
    if st.button("Mulai ➔"): move_step(1); st.rerun()

elif st.session_state.step == 1:
    st.title("⚖️ Identitas & Petunjuk")
    st.session_state.p_nama = st.text_input("Nama Panelis", st.session_state.p_nama)
    st.session_state.p_kerja = st.text_input("Pekerjaan", st.session_state.p_kerja)
    if st.button("Lanjut 🚀"):
        if st.session_state.p_nama and st.session_state.p_kerja: move_step(2); st.rerun()
        else: st.error("Lengkapi data!")

elif 2 <= st.session_state.step <= 4:
    idx_map = {2: "Pemaafan Diri", 3: "Pemaafan Orang Lain", 4: "Pemaafan Situasi"}
    aspek_aktif = idx_map[st.session_state.step]
    st.subheader(f"Bagian: {aspek_aktif}")
    
    for ind_name, items in data_aspek[aspek_aktif]:
        st.markdown(f"<div class='indicator-header'>{ind_name}</div>", unsafe_allow_html=True)
        for txt in items:
            if txt not in st.session_state.master_data: st.session_state.master_data[txt] = {"kj": 0, "rel": 0, "kes": 0, "ket": ""}
            
            # CEK APAKAH ITEM SUDAH SELESAI (Visual Feedback)
            d = st.session_state.master_data[txt]
            is_done = d["kj"] > 0 and d["rel"] > 0 and d["kes"] > 0
            card_class = "white-card completed-card" if is_done else "white-card"
            
            st.markdown(f"<div class='{card_class}'>", unsafe_allow_html=True)
            st.write(f"**{txt}**")
            c1, c2, c3 = st.columns(3)
            with c1: st.session_state.master_data[txt]["kj"] = st.selectbox("Kejelasan", [0,1,2,3,4], index=d["kj"], key=f"kj_{txt}")
            with c2: st.session_state.master_data[txt]["rel"] = st.selectbox("Relevansi", [0,1,2,3,4], index=d["rel"], key=f"rel_{txt}")
            with c3: st.session_state.master_data[txt]["kes"] = st.selectbox("Kesesuaian", [0,1,2,3,4], index=d["kes"], key=f"kes_{txt}")
            st.session_state.master_data[txt]["ket"] = st.text_input("Keterangan:", value=d["ket"], key=f"ket_{txt}")
            st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.step == 4:
        st.session_state.s_global = st.text_area("Catatan Akhir:", st.session_state.s_global)
        st.session_state.confirm_send = st.checkbox("Saya menyatakan bahwa data ini telah divalidasi dengan sebenar-benarnya.")

    nav1, nav2 = st.columns(2)
    with nav1:
        if st.button("⬅️ Kembali"): move_step(st.session_state.step - 1); st.rerun()
    with nav2:
        if st.button("🚀 KIRIM" if st.session_state.step == 4 else "Lanjut ➡️"):
            if st.session_state.step == 4 and not st.session_state.confirm_send: st.warning("Centang konfirmasi dulu.")
            else: move_step(st.session_state.step + 1); st.rerun()

elif st.session_state.step == 5:
    st.title("Finalisasi...")
    if not st.session_state.submitted:
        with st.spinner("Mengirim ke database & panel ahli..."):
            try:
                simpan_ke_gsheets()
                ex_buf = proses_excel_cvi()
                # Proses Word (Logika tetap sama seperti sebelumnya)
                doc = Document("Form Validasi Expert Judgement Ayinn Ver. 3.docx")
                # ... (Logika pengisian Word-mu)
                word_buf = io.BytesIO(); doc.save(word_buf); word_buf.seek(0)
                
                kirim_telegram_multi(word_buf, ex_buf, st.session_state.p_nama)
                st.session_state.submitted = True
                move_step(6); st.rerun()
            except Exception as e: st.error(f"Gagal: {e}")

elif st.session_state.step == 6:
    st.balloons()
    st.success("Tuntas! Data berhasil dikirim.")
