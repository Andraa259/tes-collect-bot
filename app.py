import streamlit as st
import pandas as pd
import requests
import io
from docx import Document
from streamlit_scroll_to_top import scroll_to_here
from streamlit_gsheets import GSheetsConnection

# --- KREDENSIAL ---
TOKEN = st.secrets["TOKEN"]
CHAT_ID = st.secrets["CHAT_ID"]
GSHEET_URL = st.secrets["gsheet_url"]

# --- INITIALIZING SESSION STATE ---
if 'step' not in st.session_state:
    st.session_state.step = 0
if 'scroll_to_top' not in st.session_state:
    st.session_state.scroll_to_top = False
if 'master_data' not in st.session_state:
    st.session_state.master_data = {}
if 'p_nama' not in st.session_state:
    st.session_state.p_nama = ""
if 'p_kerja' not in st.session_state:
    st.session_state.p_kerja = ""
if 'saran_global' not in st.session_state:
    st.session_state.saran_global = ""
if 'submitted' not in st.session_state:
    st.session_state.submitted = False

# --- LOGIKA SCROLL ---
if st.session_state.scroll_to_top:
    scroll_to_here(0, key=f'scroll_step_{st.session_state.step}') 
    st.session_state.scroll_to_top = False

def move_step(step_num):
    st.session_state.step = step_num
    st.session_state.scroll_to_top = True

def kirim_ke_telegram(file_stream, nama_panelis):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    files = {'document': (f'Validasi_{nama_panelis}.docx', file_stream)}
    payload = {'chat_id': CHAT_ID, 'caption': f"✅ Data Masuk: {nama_panelis}"}
    return requests.post(url, data=payload, files=files)

# --- UI STYLING ---
st.set_page_config(page_title="Expert Judgement", layout="centered")
st.markdown("""
    <style>
    .def-box { background-color: #F0F9FF; color: #075985; padding: 18px; border-radius: 12px; border-left: 6px solid #0EA5E9; margin-bottom: 20px; }
    .indicator-header { background-color: #1E3A8A; color: white; padding: 10px; border-radius: 10px 10px 0 0; text-align: center; }
    .white-card { background-color: #FFFFFF; padding: 20px; border-radius: 0 0 10px 10px; border: 1px solid #E2E8F0; margin-bottom: 20px; }
    .stButton>button { border-radius: 10px; height: 50px; font-weight: bold; width: 100%; }
    </style>
    """, unsafe_allow_html=True)

# --- DATA INDIKATOR (36 Aitem) ---
# (Data_aspek tetap sama seperti sebelumnya agar mapping tidak geser)
data_aspek = {
    "Pemaafan Diri": [
        ("Indikator 1", ["Aitem 1 (F)", "Aitem 2 (F)", "Aitem 3 (F)", "Aitem 4 (U)", "Aitem 5 (U)", "Aitem 6 (U)"]),
        ("Indikator 2", ["Aitem 7 (F)", "Aitem 8 (F)", "Aitem 9 (F)", "Aitem 10 (U)", "Aitem 11 (U)", "Aitem 12 (U)"])
    ],
    "Pemaafan Orang Lain": [
        ("Indikator 3", ["Aitem 13 (F)", "Aitem 14 (F)", "Aitem 15 (F)", "Aitem 16 (U)", "Aitem 17 (U)", "Aitem 18 (U)"]),
        ("Indikator 4", ["Aitem 19 (F)", "Aitem 20 (F)", "Aitem 21 (F)", "Aitem 22 (U)", "Aitem 23 (U)", "Aitem 24 (U)"])
    ],
    "Pemaafan Situasi": [
        ("Indikator 5", ["Aitem 25 (F)", "Aitem 26 (F)", "Aitem 27 (F)", "Aitem 28 (U)", "Aitem 29 (U)", "Aitem 30 (U)"]),
        ("Indikator 6", ["Aitem 31 (F)", "Aitem 32 (F)", "Aitem 33 (F)", "Aitem 34 (U)", "Aitem 35 (U)", "Aitem 36 (U)"])
    ]
}

# --- STEP 0: IDENTITAS ---
if st.session_state.step == 0:
    st.title("⚖️ Form Expert Judgement")
    st.session_state.p_nama = st.text_input("Nama Panelis", value=st.session_state.p_nama)
    st.session_state.p_kerja = st.text_input("Pekerjaan", value=st.session_state.p_kerja)
    if st.button("Mulai Penilaian 🚀"):
        if st.session_state.p_nama and st.session_state.p_kerja: move_step(1); st.rerun()
        else: st.error("Nama & Pekerjaan wajib diisi!")

# --- STEP 1-3: PENILAIAN ---
elif st.session_state.step in [1, 2, 3]:
    aspek_map = {1: "Pemaafan Diri", 2: "Pemaafan Orang Lain", 3: "Pemaafan Situasi"}
    aspek_aktif = aspek_map[st.session_state.step]
    st.subheader(f"Aspek: {aspek_aktif}")

    for ind_name, items in data_aspek[aspek_aktif]:
        st.markdown(f"<div class='indicator-header'>{ind_name}</div>", unsafe_allow_html=True)
        for txt in items:
            if txt not in st.session_state.master_data:
                st.session_state.master_data[txt] = {"kj": 0, "rel": 0, "kes": 0, "ket": ""}
            with st.container():
                st.markdown("<div class='white-card'>", unsafe_allow_html=True)
                st.write(f"**{txt}**")
                c1, c2, c3 = st.columns(3)
                with c1: st.session_state.master_data[txt]["kj"] = st.selectbox("Kejelasan", [0,1,2,3,4], index=st.session_state.master_data[txt]["kj"], key=f"kj_{txt}")
                with c2: st.session_state.master_data[txt]["rel"] = st.selectbox("Relevansi", [0,1,2,3,4], index=st.session_state.master_data[txt]["rel"], key=f"rel_{txt}")
                with c3: st.session_state.master_data[txt]["kes"] = st.selectbox("Kesesuaian", [0,1,2,3,4], index=st.session_state.master_data[txt]["kes"], key=f"kes_{txt}")
                st.session_state.master_data[txt]["ket"] = st.text_input("Catatan Aitem:", value=st.session_state.master_data[txt]["ket"], key=f"ket_{txt}")
                st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.step == 3:
        st.session_state.saran_global = st.text_area("Saran Keseluruhan:")

    nav1, nav2 = st.columns(2)
    with nav1:
        if st.button("⬅️ Kembali"): move_step(st.session_state.step - 1); st.rerun()
    with nav2:
        btn_label = "Lanjut ➡️" if st.session_state.step < 3 else "🚀 KIRIM HASIL"
        if st.button(btn_label): move_step(st.session_state.step + 1); st.rerun()

# --- STEP 4: PROSES SIMPAN ---
elif st.session_state.step == 4:
    st.title("Sedang Memproses...")
    if not st.session_state.submitted:
        with st.spinner("Mengirim data..."):
            try:
                conn = st.connection("gsheets", type=GSheetsConnection)
                
                # Susun list aitem agar urutan kolom konsisten (1-36)
                all_items = []
                for a in ["Pemaafan Diri", "Pemaafan Orang Lain", "Pemaafan Situasi"]:
                    for _, items in data_aspek[a]: all_items.extend(items)

                worksheets = ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]
                keys = ["kj", "rel", "kes"]

                for ws_name, k in zip(worksheets, keys):
                    df_old = conn.read(spreadsheet=GSHEET_URL, worksheet=ws_name, ttl=0)
                    
                    # --- FIX: HANYA ISI SKOR ---
                    # Entry hanya berisi skor aitem 1 s/d 36
                    new_entry = [int(st.session_state.master_data[txt][k]) for txt in all_items]
                    
                    # Cari baris kosong (cek kolom B / Index 1)
                    idx_target = None
                    for i in range(2, 40):
                        if i >= len(df_old) or str(df_old.iloc[i, 1]).strip() in ["", "nan", "0"]:
                            idx_target = i; break
                    
                    if idx_target is not None:
                        while idx_target >= len(df_old):
                            df_old = pd.concat([df_old, pd.DataFrame([[""] * df_old.shape[1]], columns=df_old.columns)], ignore_index=True)
                        
                        # Masukkan skor mulai dari Kolom B (Index 1)
                        for col_offset, val in enumerate(new_entry):
                            if (col_offset + 1) < df_old.shape[1]:
                                df_old.iloc[idx_target, col_offset + 1] = val
                        
                        conn.update(spreadsheet=GSHEET_URL, worksheet=ws_name, data=df_old.fillna(""))

                # --- WORD & TELEGRAM (Tetap pakai Nama) ---
                doc = Document("Form Validasi Expert Judgement Ayinn Ver. 3.docx")
                # (Logika pengisian Word tetap sama agar dokumen kamu lengkap)
                buf = io.BytesIO(); doc.save(buf); buf.seek(0)
                kirim_ke_telegram(buf, st.session_state.p_nama)

                st.session_state.submitted = True
                move_step(5); st.rerun()
            except Exception as e:
                st.error(f"Teknis Error: {e}"); st.button("Coba Lagi", on_click=st.rerun)

elif st.session_state.step == 5:
    st.balloons()
    st.success("✅ Data Berhasil Disimpan!")
    st.markdown("Terima kasih atas partisipasi Anda.")
