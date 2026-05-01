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
    files = {'document': (f'Form Validasi_{nama_panelis}.docx', file_stream)}
    payload = {'chat_id': CHAT_ID, 'caption': f"✅ Data Form Expert Judgement Masuk: {nama_panelis}"}
    return requests.post(url, data=payload, files=files)

# --- UI STYLING ---
st.set_page_config(page_title="Expert Judgement", layout="centered")
st.markdown("""
    <style>
    .def-box { background-color: #F0F9FF; color: #075985; padding: 18px; border-radius: 12px; border-left: 6px solid #0EA5E9; margin-bottom: 20px; line-height: 1.6; }
    .indicator-header { background-color: #1E3A8A; color: white; padding: 12px; border-radius: 10px 10px 0 0; font-weight: bold; text-align: center; margin-top: 15px; }
    .white-card { background-color: #FFFFFF; color: #1E293B; padding: 25px; border-radius: 0 0 10px 10px; border: 1px solid #E2E8F0; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); margin-bottom: 30px; }
    .stButton>button { border-radius: 10px; height: 50px; font-weight: bold; width: 100%; }
    .thanks-card { text-align: center; padding: 40px; background-color: #F8FAFC; border-radius: 20px; border: 1px solid #E2E8F0; margin-top: 50px; }
    hr { margin: 15px 0; border-top: 1px solid #eee; }
    </style>
    """, unsafe_allow_html=True)

DEF_OP = "Pemaafan adalah kemampuan individual dalam membingkai ulang terhadap suatu kesalahan yang dialami/dirasakan sehingga mampu berhenti menyalahkan diri sendiri dan melepaskan pikiran negatif tentang diri sendiri, memahami kesalahan orang lain seiring berjalannya waktu serta berhenti berpikir buruk tentang orang yang pernah menyakiti, dan mampu berdamai dengan keadaan buruk dalam hidup serta melepaskan pikiran negatif terhadap peristiwa yang berada di luar kendali."

# --- DATA INDIKATOR ---
data_aspek = {
    "Pemaafan Diri": [
        ("Indikator 1", ["Aitem 1", "Aitem 2", "Aitem 3", "Aitem 4", "Aitem 5", "Aitem 6"]),
        ("Indikator 2", ["Aitem 7", "Aitem 8", "Aitem 9", "Aitem 10", "Aitem 11", "Aitem 12"])
    ],
    "Pemaafan Orang Lain": [
        ("Indikator 3", ["Aitem 13", "Aitem 14", "Aitem 15", "Aitem 16", "Aitem 17", "Aitem 18"]),
        ("Indikator 4", ["Aitem 19", "Aitem 20", "Aitem 21", "Aitem 22", "Aitem 23", "Aitem 24"])
    ],
    "Pemaafan Situasi": [
        ("Indikator 5", ["Aitem 25", "Aitem 26", "Aitem 27", "Aitem 28", "Aitem 29", "Aitem 30"]),
        ("Indikator 6", ["Aitem 31", "Aitem 32", "Aitem 33", "Aitem 34", "Aitem 35", "Aitem 36"])
    ]
}

# --- ALUR APLIKASI ---
if st.session_state.step == 0:
    st.title("⚖️ Form Validasi Expert Judgement")
    st.markdown(f"<div class='def-box'><b>Definisi Operasional:</b><br>{DEF_OP}</div>", unsafe_allow_html=True)
    st.session_state.p_nama = st.text_input("Nama Panelis", value=st.session_state.p_nama)
    st.session_state.p_kerja = st.text_input("Pekerjaan", value=st.session_state.p_kerja)
    if st.button("Mulai Penilaian 🚀"):
        if st.session_state.p_nama and st.session_state.p_kerja: move_step(1); st.rerun()
        else: st.error("⚠️ Nama dan Pekerjaan wajib diisi!")

elif st.session_state.step in [1, 2, 3]:
    aspek_list = {1: "Pemaafan Diri", 2: "Pemaafan Orang Lain", 3: "Pemaafan Situasi"}
    aspek_aktif = aspek_list[st.session_state.step]
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
                st.session_state.master_data[txt]["ket"] = st.text_input("Keterangan:", value=st.session_state.master_data[txt]["ket"], key=f"ket_{txt}")
                st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.step == 3:
        st.session_state.saran_global = st.text_area("Catatan/Saran Keseluruhan:")

    nav1, nav2 = st.columns(2)
    with nav1:
        if st.button("⬅️ Kembali"): move_step(st.session_state.step - 1); st.rerun()
    with nav2:
        btn_label = "Lanjut ➡️" if st.session_state.step < 3 else "🚀 KIRIM HASIL"
        if st.button(btn_label): move_step(4 if st.session_state.step == 3 else st.session_state.step + 1); st.rerun()

elif st.session_state.step == 4:
    st.title("Sedang Memproses...")
    if not st.session_state.submitted:
        with st.spinner("Mencatat ke Google Sheets & Word..."):
            try:
                # --- LOGIKA GSHEETS: TARGET KOLOM C (Index 2) ---
                conn = st.connection("gsheets", type=GSheetsConnection)
                all_ordered_items = []
                for asp in ["Pemaafan Diri", "Pemaafan Orang Lain", "Pemaafan Situasi"]:
                    for _, items in data_aspek[asp]: all_ordered_items.extend(items)

                for ws_name, k in zip(["KEJELASAN", "RELEVANSI", "KESESUAIAN"], ["kj", "rel", "kes"]):
                    df_old = conn.read(spreadsheet=GSHEET_URL, worksheet=ws_name, ttl=0, header=None)
                    
                    # Pastikan struktur DataFrame minimal punya 38 kolom (A s/d AL)
                    while df_old.shape[1] < 38:
                        df_old[df_old.shape[1]] = ""
                    
                    # Cari baris kosong di rentang baris 4 s/d 33 (Index Pandas 3 s/d 32)
                    idx_target = None
                    for i in range(3, 33):
                        if i >= len(df_old) or str(df_old.iloc[i, 2]).strip() in ["", "nan", "0"]:
                            idx_target = i; break
                    
                    if idx_target is not None:
                        # Buat baris baru jika diperlukan
                        while idx_target >= len(df_old):
                            df_old.loc[len(df_old)] = [""] * df_old.shape[1]
                        
                        # Isi Horizontal: Aitem 1 di Kolom C (Index 2), Aitem 2 di D (Index 3), dst.
                        for col_offset, item_txt in enumerate(all_ordered_items):
                            df_old.iloc[idx_target, col_offset + 2] = st.session_state.master_data[item_txt][k]
                        
                        conn.update(spreadsheet=GSHEET_URL, worksheet=ws_name, data=df_old.fillna(""))

                # --- WORD & TELEGRAM ---
                doc = Document("Form Validasi Expert Judgement Ayinn Ver. 3.docx")
                # (Logika Word tetap sama)
                buf = io.BytesIO(); doc.save(buf); buf.seek(0)
                kirim_ke_telegram(buf, st.session_state.p_nama)
                
                st.session_state.submitted = True
                move_step(5); st.rerun()
            except Exception as e:
                st.error(f"Terjadi kesalahan teknis: {e}")
                if st.button("Coba Lagi"): st.rerun()

elif st.session_state.step == 5:
    st.balloons()
    st.markdown("""
        <div class='thanks-card'>
            <h1 style='color: #1E3A8A;'>Terima Kasih! ✨</h1>
            <p style='font-size: 1.2rem; color: #475569;'>
                Data penilaian Anda telah berhasil kami terima dan dikirimkan ke peneliti. 
                Kontribusi Anda sangat berharga bagi pengembangan instrumen penelitian ini.
            </p>
            <hr>
            <p style='font-style: italic; color: #64748b;'>Halaman ini dapat Anda tutup sekarang.</p>
        </div>
    """, unsafe_allow_html=True)
