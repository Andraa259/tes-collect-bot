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
    payload = {'chat_id': CHAT_ID, 'caption': f"✅ Data Expert Judgement Masuk: {nama_panelis}"}
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
        ("Indikator 1: Kemampuan untuk berhenti menyalahkan diri sendiri", [
            "Seiring waktu, saya bisa memaklumi kesalahan pribadi yang pernah dilakukan. (Favorable)",
            "Ketika membuat kesalahan, saya fokus pada perbaikan daripada terus menerus menyalahkan diri sendiri. (Favorable)",
            "Saya memilih untuk berdamai dengan kekurangan diri sendiri. (Favorable)",
            "Sulit bagi saya untuk berhenti menyalahkan diri sendiri. (Unfavorable)",
            "Muncul perasaan benci ketika saya mengingat kesalahan diri sendiri. (Unfavorable)",
            "Saya terjebak dalam penyesalan atas kegagalan diri sendiri. (Unfavorable)"
        ]),
        ("Indikator 2: Kesediaan untuk melepaskan pikiran negatif tentang diri", [
            "Pikiran negatif tentang diri sendiri mulai memudar seiring waktu. (Favorable)",
            "Saya dapat memahami diri sendiri atas kesalahan yang telah saya lakukan. (Favorable)",
            "Saat ingatan yang mengganggu tentang diri sendiri muncul, saya mampu melepaskannya. (Favorable)",
            "Sulit bagi saya untuk berhenti memikirkan hal-hal buruk yang pernah menimpa diri sendiri. (Unfavorable)",
            "Pikiran tentang kesalahan diri sendiri terus muncul walaupun sudah berusaha melupakannya. (Unfavorable)",
            "Saya sering susah berkonsentrasi karena teringat pada kesalahan diri sendiri yang telah lalu. (Unfavorable)"
        ])
    ],
    "Pemaafan Orang Lain": [
        ("Indikator 3: Kemampuan untuk memahami kesalahan orang lain", [
            "Saya dapat memaklumi bahwa setiap orang pasti pernah melakukan kekeliruan. (Favorable)",
            "Saya mencoba memahami alasan dibalik tindakan orang lain yang telah menyakiti saya. (Favorable)",
            "Saya menyadari bahwa ada alasan tertentu yang membuat orang lain sulit untuk bertindak benar. (Favorable)",
            "Memandang orang yang menyakiti saya sebagai pribadi yang memiliki karakter buruk. (Unfavorable)",
            "Saya tidak bisa menerima alasan apapun dari orang yang telah mengecewakan saya. (Unfavorable)",
            "Sangat sulit bagi saya untuk mengerti mengapa seseorang berbuat jahat kepada saya. (Unfavorable)"
        ]),
        ("Indikator 4: Berhenti berpikir buruk tentang orang yang pernah menyakiti", [
            "Pikiran buruk terhadap orang yang pernah menyakiti saya perlahan mulai menghilang. (Favorable)",
            "Saya merasa sudah tidak lagi menyimpan kebencian terhadap orang yang pernah menyakiti saya. (Favorable)",
            "Mudah bagi saya melepaskan rasa benci yang tertuju pada orang yang pernah berbuat salah. (Favorable)",
            "Saya terus membayangkan hal-hal negatif terjadi pada orang yang telah menyakiti saya. (Unfavorable)",
            "Sulit bagi saya untuk menghilangkan pandangan negatif terhadap orang yang pernah berbuat salah. (Unfavorable)",
            "Rasa kesal muncul kembali setiap kali saya mengingat perlakuan orang yang menyakiti saya. (Unfavorable)"
        ])
    ],
    "Pemaafan Situasi": [
        ("Indikator 5: Kemampuan untuk berdamai dengan keadaan buruk dalam hidup", [
            "Seiring berjalannya waktu, saya mulai bisa menerima kenyataan pahit yang terjadi dalam hidup dengan lapang dada. (Favorable)",
            "Saya sadar untuk tidak menyalahkan nasib atas kejadian buruk yang menimpa. (Favorable)",
            "Mampu menerima kenyataan bahwa hidup tidak selalu berjalan sesuai dengan rencana saya. (Favorable)",
            "Saya merasa semesta tidak adil karena terus memberikan cobaan yang berat. (Unfavorable)",
            "Sering merasa terjebak dalam nasib buruk yang seolah-olah tidak pernah berakhir di hidup saya. (Unfavorable)",
            "Terus-menerus mengeluhkan nasib buruk yang menimpa diri saya menjadi hal yang sulit untuk dihentikan. (Unfavorable)"
        ]),
        ("Indikator 6: Melepaskan pikiran negatif terhadap peristiwa luar kendali", [
            "Pikiran tentang kejadian buruk di masa lalu tidak lagi mengganggu saya untuk berkonsentrasi sehari-hari. (Favorable)",
            "Saya merasa sudah bisa berdamai dengan bayangan tentang masa-masa sulit yang pernah dialami. (Favorable)",
            "Saya mampu mengalihkan fokus dari peristiwa yang mengecewakan ke hal-hal yang lebih produktif. (Favorable)",
            "Sangat sulit bagi saya untuk tidak memikirkan kegagalan yang pernah dialami. (Unfavorable)",
            "Saya merasa terjebak dalam memori tentang kejadian buruk yang pernah saya alami. (Unfavorable)",
            "Bayangan mengenai ketidakadilan hidup di masa lalu sering kali muncul tanpa bisa saya kendalikan. (Unfavorable)"
        ])
    ]
}

# --- ALUR APLIKASI ---

if st.session_state.step == 0:
    st.title("⚖️ Form Validasi Expert Judgement")
    st.markdown(f"<div class='def-box'><b>Definisi Operasional:</b><br>{DEF_OP}</div>", unsafe_allow_html=True)
    st.subheader("📝 PETUNJUK PENGISIAN")
    st.markdown("""
    * **Kejelasan**: Bahasa mudah dipahami.
    * **Relevansi**: Aitem menggambarkan variabel.
    * **Kesesuaian**: Aitem sesuai indikator.
    * Skor **1-4** (1: Kurang, 4: Baik Sekali). Skor **0** berarti belum diisi.
    """)
    
    st.session_state.p_nama = st.text_input("Nama Panelis", value=st.session_state.p_nama)
    st.session_state.p_kerja = st.text_input("Pekerjaan", value=st.session_state.p_kerja)
    
    if st.button("Mulai Penilaian 🚀"):
        if st.session_state.p_nama == "" or st.session_state.p_kerja == "":
            st.error("⚠️ Nama dan Pekerjaan wajib diisi!")
        else: move_step(1); st.rerun()

elif st.session_state.step in [1, 2, 3]:
    aspek_list = {1: "Pemaafan Diri", 2: "Pemaafan Orang Lain", 3: "Pemaafan Situasi"}
    aspek_aktif = aspek_list[st.session_state.step]
    st.subheader(f"Aspek: {aspek_aktif}")

    current_page_items = []
    for _, items in data_aspek[aspek_aktif]: current_page_items.extend(items)

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
                st.session_state.master_data[txt]["ket"] = st.text_input("Keterangan Aitem:", value=st.session_state.master_data[txt]["ket"], key=f"ket_{txt}")
                st.markdown("</div>", unsafe_allow_html=True)

    errors = [t for t in current_page_items if any(st.session_state.master_data[t][k] == 0 for k in ["kj", "rel", "kes"])]

    if st.session_state.step == 3:
        st.session_state.saran_global = st.text_area("Saran Keseluruhan:", value=st.session_state.saran_global)

    nav1, nav2 = st.columns(2)
    with nav1:
        if st.button("⬅️ Kembali"): move_step(st.session_state.step - 1); st.rerun()
    with nav2:
        btn_label = "Lanjut ➡️" if st.session_state.step < 3 else "🚀 KIRIM HASIL"
        if st.button(btn_label):
            if False: st.error(f"⚠️ Lengkapi {len(errors)} aitem di halaman ini (skor tidak boleh 0).")
            else: move_step(4 if st.session_state.step == 3 else st.session_state.step + 1); st.rerun()

elif st.session_state.step == 4:
    st.title("Sedang Memproses...")
    if not st.session_state.submitted:
        with st.spinner("Mencatat data ke Excel, Word, & Telegram..."):
            try:
                # --- 1. KONEKSI GSHEETS ---
                conn = st.connection("gsheets", type=GSheetsConnection)
                all_ordered_items = []
                for aspect in ["Pemaafan Diri", "Pemaafan Orang Lain", "Pemaafan Situasi"]:
                    for _, items in data_aspek[aspect]: all_ordered_items.extend(items)

                worksheets = ["KEJELASAN", "RELEVANSI", "KESESUAIAN"]
                keys = ["kj", "rel", "kes"]

                for ws_name, k in zip(worksheets, keys):
                    df_old = conn.read(spreadsheet=GSHEET_URL, worksheet=ws_name, ttl=0)
                    
                    # Skor murni (Aitem 1-36) sebagai list integer
                    new_entry_scores = [int(st.session_state.master_data[txt][k]) for txt in all_ordered_items]
                    
                    # Logika Cari Baris Kosong (Cek Kolom B / Index 1)
                    idx_target = None
                    for i in range(2, 50): # Cek dari baris 3 ke bawah
                        if i >= len(df_old):
                            idx_target = i; break
                        val_check = str(df_old.iloc[i, 1]).strip().lower()
                        if val_check in ["", "nan", "0", "0.0", "none"]:
                            idx_target = i; break
                    
                    if idx_target is not None:
                        while idx_target >= len(df_old):
                            df_old = pd.concat([df_old, pd.DataFrame([[""] * df_old.shape[1]], columns=df_old.columns)], ignore_index=True)
                        
                        # ISI SKOR MULAI KOLOM B (INDEX 1). KOLOM A (INDEX 0) DILEWATI.
                        for col_offset, val in enumerate(new_entry_scores):
                            target_col = col_offset + 1 
                            if target_col < df_old.shape[1]:
                                df_old.iloc[idx_target, target_col] = val
                        
                        conn.update(spreadsheet=GSHEET_URL, worksheet=ws_name, data=df_old.fillna(""))
                    else:
                        st.warning(f"Sheet {ws_name} penuh.")

                # --- 2. PROSES WORD ---
                doc = Document("Form Validasi Expert Judgement Ayinn Ver. 3.docx")
                for p in doc.paragraphs:
                    if "Nama\t\t:" in p.text: p.text = f"Nama\t\t: {st.session_state.p_nama}"
                    if "Pekerjaan\t:" in p.text: p.text = f"Pekerjaan\t: {st.session_state.p_kerja}"
                
                table = doc.tables[0]
                for row in table.rows:
                    aitem_word = "".join(row.cells[2].text.split()).lower()
                    for txt_ori, data in st.session_state.master_data.items():
                        if "".join(txt_ori.split()).lower()[:50] in aitem_word:
                            row.cells[3].text, row.cells[4].text = str(data["kj"]), str(data["rel"])
                            row.cells[5].text, row.cells[6].text = str(data["kes"]), str(data["ket"])
                
                for row in table.rows:
                    if "Catatan" in row.cells[2].text: row.cells[2].text += f"\n{st.session_state.saran_global}"

                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)

                # --- 3. TELEGRAM ---
                kirim_ke_telegram(buf, st.session_state.p_nama)

                st.session_state.submitted = True
                move_step(5); st.rerun()

            except Exception as e:
                st.error(f"Terjadi kesalahan teknis: {e}")
                if st.button("Coba Lagi"): st.rerun()

elif st.session_state.step == 5:
    st.balloons()
    st.markdown("<div class='thanks-card'><h1>Terima Kasih! ✨</h1><p>Data Anda telah berhasil terkirim.</p></div>", unsafe_allow_html=True)
