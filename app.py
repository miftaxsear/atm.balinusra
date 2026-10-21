import json
import os
import glob
import sqlite3
import numpy as np
import pandas as pd
import streamlit as st

# 1. Konfigurasi Halaman & CSS Custom untuk Lebar Tombol & Filter agar Tidak Terlalu Memanjang ke Samping
st.set_page_config(page_title="Balinusra Monitoring", layout="wide")

st.markdown(
    """
    <style>
        [data-testid="stSidebarNav"] {
            display: none;
        }

        .block-container {
            padding-top: 2.5rem !important;
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
            max-width: 100% !important;
        }

        /* --- CSS CUSTOM UNTUK MEMPERKECIL TOMBOL & FILTER AGAR TIDAK MELEBAR --- */
        div.stButton > button {
            max-width: 240px !important;
            min-height: 32px !important;
            padding-top: 2px !important;
            padding-bottom: 2px !important;
            font-size: 13px !important;
        }
        div.stSelectbox > div > div {
            min-height: 32px !important;
            padding-top: 2px !important;
            padding-bottom: 2px !important;
            font-size: 13px !important;
        }
        [data-testid="column"] {
            padding: 0px 4px !important;
        }
        /* ---------------------------------------------------------------------- */

        div[data-testid="stSelectbox"]:has(label:contains("Pilih Tahun:")) label p,
        div[data-testid="stSelectbox"]:has(label:contains("Pilih Bulan:")) label p,
        div[data-testid="stSelectbox"]:has(label:contains("Pilih Bulan Standby:")) label p {
            background-color: #AFEEEE !important;
            color: #004D40 !important;
            padding: 2px 8px !important;
            border-radius: 4px !important;
            font-weight: bold !important;
            display: inline-block !important;
        }

        div[data-testid="stSelectbox"]:has(label:contains("Pilih / Ketik WSID:")) label p,
        div[data-testid="stSelectbox"]:has(label:contains("Pilih CSE / Pengelola:")) label p,
        div[data-testid="stSelectbox"]:has(label:contains("Pilih / Ketik CSE:")) label p,
        div[data-testid="stSelectbox"]:has(label:contains("Pilih / Ketik PKT:")) label p {
            background-color: #FFB6C1 !important;
            color: #880E4F !important;
            padding: 2px 8px !important;
            border-radius: 4px !important;
            font-weight: bold !important;
            display: inline-block !important;
        }

        /* --- TAMBAHAN CSS RESPONSIF UNTUK PERANGKAT MOBILE / HP --- */
        @media screen and (max-width: 768px) {
            .stColumns {
                flex-direction: row !important;
                flex-wrap: wrap;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# =============================================================
# DETEKSI OTOMATIS TAHUN & BULAN DARI MASTER DATA TERAKHIR
# =============================================================
excel_files = glob.glob("MASTER_*.xlsx") + glob.glob("master_*.xlsx") + glob.glob("*.xlsx")

detected_year = 2026
detected_month_num = 10  # Oktober (berdasarkan data master terakhir)

if excel_files:
    latest_file = max(excel_files, key=os.path.getmtime)
    filename_only = os.path.basename(latest_file)
    try:
        parts = filename_only.replace(".xlsx", "").split("_")
        for p in parts:
            if "-" in p and len(p) == 7:  # Format YYYY-MM
                y_str, m_str = p.split("-")
                detected_year = int(y_str)
                detected_month_num = int(m_str)
                break
    except:
        pass

# Variabel Global untuk Tahun & Bulan Aktif (otomatis merujuk ke data master terakhir)
selected_tahun = detected_year
mapping_angka_ke_bulan = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr",
    5: "Mei", 6: "Jun", 7: "Jul", 8: "Agu",
    9: "Sep", 10: "Okt", 11: "Nov", 12: "Des"
}
selected_bulan_nama = mapping_angka_ke_bulan.get(int(detected_month_num), "Okt")
selected_month = f"{selected_tahun}-{str(detected_month_num).zfill(2)}"

import pandas as pd
import sqlite3

# =============================================================
# PANEL KONTROL (Hanya Tampil untuk Admin)
# =============================================================
if st.session_state.get("role") == "admin":
    st.sidebar.subheader("Panel Kontrol")
    
    if st.sidebar.button("🔄 Refresh & Update Database.db", key="btn_refresh_db", use_container_width=True):
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() else os.getcwd()
            pattern = os.path.join(base_dir, "MASTER_*.xlsx")
            matching_files = glob.glob(pattern)
            
            if not matching_files:
                st.sidebar.error("❌ File MASTER_*.xlsx tidak ditemukan di folder server!")
            else:
                latest_file = max(matching_files, key=os.path.getmtime)
                xls = pd.ExcelFile(latest_file)
                db_path = os.path.join(base_dir, "MASTER_2026-10_db.db")
                conn = sqlite3.connect(db_path)
                
                for sheet_name in xls.sheet_names:
                    try:
                        if sheet_name.isdigit() and 1 <= int(sheet_name) <= 31:
                            table_name = f"day_{sheet_name}"
                            df_sheet = pd.read_excel(xls, sheet_name=sheet_name, header=5)
                        else:
                            table_name = sheet_name.strip().lower().replace(' ', '_').replace('&', 'and')
                            df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
                        
                        df_sheet.columns = [str(col).strip() for col in df_sheet.columns]
                        df_sheet.to_sql(table_name, conn, if_exists="replace", index=False)
                    except Exception:
                        continue
                
                conn.close()
                st.sidebar.success(f"✅ Sukses Update Data s/d Tanggal 19!")
                st.balloons()
                st.rerun()
                
        except Exception as e:
            st.sidebar.error(f"❌ Error: {e}")
            
    st.sidebar.markdown("---")

# =============================================================
# FUNGSI INISIALISASI & MIGRASI OTOMATIS KE SQLITE (.db)
# =============================================================
def initialize_database_from_excel(excel_file_path, year_str, m_num):
    db_name = "atm_database.db"
    month_code = str(int(m_num))
    table_teritori = f"teritori_mesin_{year_str}_{month_code}"
    
    try:
        conn = sqlite3.connect(db_name)
        cursor = conn.cursor()
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_teritori}';")
        tbl_exists = cursor.fetchone()
        
        if not tbl_exists and excel_file_path and os.path.exists(excel_file_path):
            df_teritori = pd.read_excel(excel_file_path, sheet_name=0) 
            df_teritori.columns = [str(c).strip().upper().replace(" ", "_") for c in df_teritori.columns]
            df_teritori.to_sql(table_teritori, conn, if_exists='replace', index=False)
            
        conn.close()
    except Exception as e:
        print(f"Gagal menginisialisasi database SQLite: {e}")

target_excel_file = next((f for f in excel_files if f"{detected_year}-{detected_month_num:02d}" in f), None) if excel_files else None

if target_excel_file and os.path.exists(target_excel_file):
    try:
        y_val, m_val = str(detected_year), str(detected_month_num)
        initialize_database_from_excel(target_excel_file, y_val, m_val)
    except:
        pass


# =============================================================
# INISIALISASI DATABASE USER SQLite & VERIFIKASI LOGIN
# =============================================================
def init_user_db():
    conn = sqlite3.connect("app_users.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            role TEXT NOT NULL,
            name TEXT NOT NULL
        )
    """)
    
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        default_users = [
            ("admin", "balinusra1", "admin", "Administrator"),
            ("engineer", "engineer234", "viewer", "Engineer")
        ]
        cursor.executemany("INSERT OR IGNORE INTO users VALUES (?, ?, ?, ?)", default_users)
        conn.commit()
    conn.close()

init_user_db()

def verify_login(username, password):
    conn = sqlite3.connect("app_users.db")
    cursor = conn.cursor()
    cursor.execute("SELECT password, role, name FROM users WHERE username = ?", (username,))
    result = cursor.fetchone()
    conn.close()
    
    if result and result[0] == password:
        return {"success": True, "role": result[1], "name": result[2]}
    return {"success": False}


# =============================================================
# SESSION STATE & NAVIGATION SETUP (HARUS DI ATAS SEBELUM DI-CHECK)
# =============================================================
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["role"] = None
    st.session_state["user_name"] = ""

menu_options_list = [
    "Uptime FRX DT (WSID)",
    "Uptime Harian",
    "Uptime CSE by Tipe Mesin",
    "Uptime PKT by Tipe Mesin",
    "Riwayat Kunjungan (Visit)",
    "Teritori Mesin Engineer",
    "Jadwal Standby CSE",
]

if "nav_menu" not in st.session_state:
    st.session_state["nav_menu"] = "Uptime FRX DT (WSID)"

if "nav_history" not in st.session_state:
    st.session_state["nav_history"] = []


def change_page(target_menu):
    current = st.session_state["nav_menu"]
    if target_menu != current:
        st.session_state["nav_history"].append(current)
        st.session_state["nav_menu"] = target_menu
        st.rerun()


def go_back_history():
    if len(st.session_state["nav_history"]) > 0:
        prev_page = st.session_state["nav_history"].pop()
        st.session_state["nav_menu"] = prev_page
    else:
        st.session_state["nav_menu"] = "Uptime FRX DT (WSID)"
    st.rerun()


# =============================================================
# HALAMAN LOGIN STYLE MYDATINDO (Dengan Auto-Fill Saved Credentials)
# =============================================================
def show_login_page():
    login_css = """
    <style>
        header {visibility: hidden;}
        .block-container {
            padding-top: 3rem !important;
            padding-bottom: 2rem !important;
        }
        
        .brand-title {
            font-family: 'Lucida Calligraphy', 'Lucida Handwriting', 'Apple Chancery', cursive;
            color: #27ae60;
            font-size: 34px;
            font-weight: bold;
            text-align: center;
            margin-bottom: 12px;
        }

        .login-subtitle {
            text-align: center;
            color: #666666;
            font-size: 13px;
            margin-bottom: 18px;
        }

        div[data-baseweb="input"] {
            background-color: #eef4fb !important;
            border: 1px solid #b8d3f2 !important;
            border-radius: 3px !important;
        }
        
        div[data-baseweb="input"] input {
            color: #2c3e50 !important;
            font-size: 14px !important;
        }

        div.stButton > button {
            width: 100% !important;
            background-color: #3488b5 !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 3px !important;
            padding: 8px 0px !important;
            font-weight: bold !important;
            font-size: 14px !important;
            margin-top: 8px !important;
        }

        div.stButton > button:hover {
            background-color: #286f95 !important;
            color: #ffffff !important;
        }

        .login-footer {
            margin-top: 22px;
            font-size: 11px;
            color: #777777;
            text-align: center;
        }
    </style>
    """
    st.markdown(login_css, unsafe_allow_html=True)

    if "saved_username" not in st.session_state:
        st.session_state["saved_username"] = "admin"
    if "saved_password" not in st.session_state:
        st.session_state["saved_password"] = "balinusra1"

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown(
            '<div class="brand-title">Balinusra Monitoring</div>',
            unsafe_allow_html=True,
        )

        with st.container(border=True):
            st.markdown(
                '<div class="login-subtitle">Sign in to start your session</div>',
                unsafe_allow_html=True,
            )

            with st.form("login_form"):
                username_input = st.text_input(
                    "Username",
                    value=st.session_state["saved_username"],
                    placeholder="Username",
                    label_visibility="collapsed",
                )
                password_input = st.text_input(
                    "Password",
                    value=st.session_state["saved_password"],
                    type="password",
                    placeholder="Password",
                    label_visibility="collapsed",
                )
                submit_button = st.form_submit_button("Sign In")

                if submit_button:
                    auth = verify_login(username_input, password_input)
                    if auth["success"]:
                        st.session_state["saved_username"] = username_input
                        st.session_state["saved_password"] = password_input
                        
                        st.session_state["logged_in"] = True
                        st.session_state["role"] = auth["role"]
                        st.session_state["user_name"] = auth["name"]
                        st.success("Login berhasil!")
                        st.rerun()
                    else:
                        st.error("Username atau Password salah!")



            st.markdown(
                '<div class="login-footer">Copyright © 2026 <strong>MiftaxSear</strong>. All rights reserved.</div>',
                unsafe_allow_html=True,
            )


# =============================================================
# KONTROL TAMPILAN UTAMA APLIKASI
# =============================================================
if not st.session_state["logged_in"]:
    show_login_page()
    st.stop()
else:
    current_page = st.session_state["nav_menu"]

# -------------------------------------------------------------
# 4. FUNGSI MIGRASI & MEMBACA DATA MASTER KE SQLITE
# -------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_data(file_path):
    db_file = file_path.replace(".xlsx", "_db.db")
    
    if not os.path.exists(db_file) and os.path.exists(file_path):
        conn = sqlite3.connect(db_file)
        xls = pd.ExcelFile(file_path)
        
        for idx_sheet, s_name in enumerate(xls.sheet_names):
            try:
                s_str = str(s_name).strip()
                if s_str.isdigit():
                    df = pd.read_excel(xls, s_name, header=5)
                else:
                    df = pd.read_excel(xls, s_name)
                
                new_cols = []
                seen = {}
                for c in df.columns:
                    clean_c = str(c).strip().replace(" ", "_").replace(".", "_").replace("-", "_")
                    if not clean_c or clean_c.lower() == "nan":
                        clean_c = "Unnamed"
                    if clean_c in seen:
                        seen[clean_c] += 1
                        clean_c = f"{clean_c}_{seen[clean_c]}"
                    else:
                        seen[clean_c] = 0
                    new_cols.append(clean_c)
                
                df.columns = new_cols
                df = df.astype(str)
                
                if "site" in s_str.lower() or "blns" in s_str.lower():
                    table_name = "site_blns"
                elif "atm" in s_str.lower() and "total" in s_str.lower():
                    table_name = "atm_total"
                elif s_str.isdigit():
                    table_name = f"day_{s_str}"
                else:
                    table_name = s_str.lower().replace(" ", "_").replace("-", "_")
                
                df.to_sql(table_name, conn, if_exists="replace", index=False)
            except Exception:
                pass
        conn.close()

    conn = sqlite3.connect(db_file)
    tables = [row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
    
    df_frx_raw = pd.read_sql("SELECT * FROM frx_dt", conn) if "frx_dt" in tables else pd.DataFrame()
    df_site = pd.read_sql("SELECT * FROM site_blns", conn) if "site_blns" in tables else pd.DataFrame()
    df_visit = pd.read_sql("SELECT * FROM mainvisit", conn) if "mainvisit" in tables else pd.DataFrame()
    
    df_atm_summary = pd.DataFrame()
    target_atm_t = next((t for t in tables if "atm" in t and "total" in t), None)
    if target_atm_t:
        df_atm_summary = pd.read_sql(f"SELECT * FROM {target_atm_t}", conn)

    daily_sheets = {}
    for t in tables:
        if t.startswith("day_"):
            day_num = t.replace("day_", "")
            df_day = pd.read_sql(f"SELECT * FROM {t}", conn)
            if "WSID" in df_day.columns or "wsid" in df_day.columns:
                daily_sheets[int(day_num)] = df_day

    conn.close()
    return df_frx_raw, df_site, df_visit, df_atm_summary, daily_sheets

excel_files = glob.glob("MASTER_*.xlsx") + glob.glob("master_*.xlsx")

default_year = 2026
default_month_num = "10"

if excel_files:
    latest_file = max(excel_files, key=os.path.getmtime)
    filename_only = os.path.basename(latest_file)
    try:
        parts = filename_only.replace(".xlsx", "").split("_")
        for p in parts:
            if "-" in p and len(p) == 7:
                y_str, m_str = p.split("-")
                default_year = int(y_str)
                default_month_num = m_str
                break
    except:
        pass

year_options = [2024, 2025, 2026, 2027]
if default_year not in year_options:
    year_options.append(default_year)
    year_options = sorted(list(set(year_options)))

available_months_map = {}
for f in excel_files:
    base_name = os.path.basename(f)
    if base_name.startswith("MASTER_") and "-" in base_name:
        try:
            parts = base_name.replace(".xlsx", "").replace("MASTER_", "").split("-")
            if len(parts) == 2:
                y_val = int(parts[0])
                m_val = parts[1]
                if y_val not in available_months_map:
                    available_months_map[y_val] = []
                if m_val not in available_months_map[y_val]:
                    available_months_map[y_val].append(m_val)
        except:
            pass

if not available_months_map:
    available_months_map = {y: [f"{m:02d}" for m in range(1, 13)] for y in [2024, 2025, 2026, 2027]}


# -------------------------------------------------------------
# 5. SIDEBAR: NAVIGASI MENU KUSTOM & PANEL ADMIN
# -------------------------------------------------------------
st.sidebar.markdown(
    "<h3 style='font-size: 18px; font-weight: bold; margin-bottom: 10px; white-space: nowrap;'>🔍 Dasbord Monitoring</h3>",
    unsafe_allow_html=True,
)

st.sidebar.markdown(
    "<p style='font-size: 13px; font-weight: 600; color: #555; margin-bottom: 5px;'>Pilih Tampilan Dashboard:</p>",
    unsafe_allow_html=True,
)

menu_icons = {
    "Uptime FRX DT (WSID)": "🏠",
    "Uptime Harian": "📊",
    "Uptime CSE by Tipe Mesin": "👨‍🔧",
    "Uptime PKT by Tipe Mesin": "⚙️",
    "Riwayat Kunjungan (Visit)": "📋",
    "Teritori Mesin Engineer": "🗺️",
    "Jadwal Standby CSE": "📅",
}

for m_item in menu_options_list:
    is_active = st.session_state["nav_menu"] == m_item
    btn_type = "primary" if is_active else "secondary"
    icon = menu_icons.get(m_item, "📌")
    if st.sidebar.button(
        f"{icon} {m_item}",
        key=f"nav_btn_{m_item}",
        use_container_width=True,
        type=btn_type,
    ):
        change_page(m_item)

menu_option = st.session_state["nav_menu"]

st.sidebar.divider()

if st.session_state["role"] == "admin":
    st.sidebar.subheader("⚙️ Panel Admin (All Access)")
    admin_upload_year = st.sidebar.selectbox(
        "Pilih Tahun File Master:",
        options=year_options,
        index=year_options.index(default_year) if default_year in year_options else 0,
    )
    
    default_month_idx = int(default_month_num) - 1 if str(default_month_num).isdigit() and 1 <= int(default_month_num) <= 12 else 9
    
    admin_upload_month_num = st.sidebar.selectbox(
        "Pilih Bulan File Master:",
        options=[f"{m:02d}" for m in range(1, 13)],
        index=default_month_idx,
    )
    admin_upload_month = f"{admin_upload_year}-{admin_upload_month_num}"

    uploaded_file = st.sidebar.file_uploader(
        f"Upload File Master ({admin_upload_month}):", type=["xlsx"]
    )

    if uploaded_file is not None:
        save_file_name = f"MASTER_{admin_upload_month}.xlsx"
        with open(save_file_name, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.cache_data.clear()
        st.sidebar.success(f"File {save_file_name} berhasil diperbarui!")
        st.rerun()

    st.sidebar.markdown("---")
    uploaded_standby_file = st.sidebar.file_uploader(
        "Upload File Jadwal Standby:", type=["xlsx"]
    )
    if uploaded_standby_file is not None:
        file_save_name = uploaded_standby_file.name
        with open(file_save_name, "wb") as f:
            f.write(uploaded_standby_file.getbuffer())
        st.cache_data.clear()
        st.sidebar.success(f"File Jadwal Standby ({file_save_name}) berhasil diperbarui!")
        st.rerun()
else:
    st.sidebar.info("👁️ **Mode Viewer**: Anda hanya memiliki akses melihat data.")

st.sidebar.divider()

st.sidebar.markdown(f"**Logged in as:** {st.session_state['user_name']}")
st.sidebar.markdown(f"**Role:** `{st.session_state['role'].upper()}`")

if st.sidebar.button("Logout"):
    st.session_state["logged_in"] = False
    st.session_state["role"] = None
    st.rerun()


selected_year = default_year
selected_month_num = default_month_num
selected_month = f"{selected_year}-{selected_month_num}"
target_excel_file = f"MASTER_{selected_month}.xlsx"

month_names_dict = {
    "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
    "05": "Mei", "06": "Jun", "07": "Jul", "08": "Agu",
    "09": "Sep", "10": "Okt", "11": "Nov", "12": "Des"
}

def get_ordered_months(year):
    base_months = available_months_map.get(year, [f"{m:02d}" for m in range(1, 13)])
    sorted_months = sorted(list(set(base_months)), key=lambda x: int(x), reverse=True)
    if selected_month_num in sorted_months:
        sorted_months.remove(selected_month_num)
        sorted_months.insert(0, selected_month_num)
    return sorted_months

# =============================================================
# BAGIAN ATAS DASBOR: LEBAR KOLOM & FILTER DISAMAKAN RATA
# =============================================================
if menu_option not in ["Teritori Mesin Engineer", "Jadwal Standby CSE"]:
    # Mengatur bobot kolom agar tombol riwayat, tahun, bulan, dan WSID sejajar rapi ke samping
    col_btn, col_year, col_month, col_wsid, col_space_right = st.columns([1.5, 0.9, 0.9, 1.2, 1.5])

    with col_btn:
        # Menyelaraskan posisi vertikal tombol agar sejajar dengan tinggi input selectbox di sebelah kanannya
        st.markdown("<div style='margin-top: 27px;'></div>", unsafe_allow_html=True)
        # Tombol link menuju ke halaman /maps yang sudah terproteksi login
        st.link_button(f"🔗 Buka Peta ATM ({selected_wsid if 'selected_wsid' in locals() else 'Maps'})", "/maps", use_container_width=True)

    with col_year:
        selected_year = st.selectbox(
            "Pilih Tahun:",
            options=year_options,
            index=year_options.index(default_year) if default_year in year_options else 0,
            key="main_year_select",
        )
    
    with col_month:
        months_list = get_ordered_months(selected_year)
        if selected_month_num not in months_list and months_list:
            selected_month_num = months_list[0]
        
        selected_month_num = st.selectbox(
            "Pilih Bulan:",
            options=months_list,
            format_func=lambda x: month_names_dict.get(x, x),
            index=months_list.index(selected_month_num) if selected_month_num in months_list else 0,
            key="main_month_select"
        )
    
    with col_wsid:
        # Bagian pilihan WSID agar proporsional dan tidak terpotong di HP
        if "available_wsids" in locals() and available_wsids:
            selected_wsid = st.selectbox(
                "Pilih / Ketik WSID:",
                options=available_wsids,
                index=available_wsids.index(default_wsid) if default_wsid in available_wsids else 0,
                key="main_wsid_select"
            )

    selected_month = f"{selected_year}-{selected_month_num}"
    target_excel_file = f"MASTER_{selected_month}.xlsx"

    if not os.path.exists(target_excel_file):
        matching_masters = glob.glob(f"MASTER_{selected_year}-*.xlsx")
        if matching_masters:
            matching_masters.sort(reverse=True)
            target_excel_file = matching_masters[0]
            selected_month = target_excel_file.replace("MASTER_", "").replace(".xlsx", "")
        elif os.path.exists("MASTER.xlsx"):
            target_excel_file = "MASTER.xlsx"

elif menu_option == "Teritori Mesin Engineer":
    if not st.session_state.get("teritory_detail_wsid"):
        col_title, col_year, col_month, col_back = st.columns([2.4, 0.7, 0.9, 0.8])
        
        with col_title:
            st.markdown(
                f"<h4 style='margin-top: -5px; margin-bottom: 12px; color: #2c3e50; font-weight: 600; white-space: nowrap;'>🗺️ Teritori Mesin Engineer ({selected_month})</h4>",
                unsafe_allow_html=True,
            )

        with col_year:
            selected_year = st.selectbox(
                "Pilih Tahun:",
                options=year_options,
                index=year_options.index(default_year) if default_year in year_options else 0,
                key="teritory_year_select",
            )

        with col_month:
            months_list = get_ordered_months(selected_year)
            if selected_month_num not in months_list and months_list:
                selected_month_num = months_list[0]
            
            selected_month_num = st.selectbox(
                "Pilih Bulan:",
                options=months_list,
                format_func=lambda x: month_names_dict.get(x, x),
                index=months_list.index(selected_month_num) if selected_month_num in months_list else 0,
                key="teritory_month_select"
            )

        with col_back:
            st.write("")
            st.button("⬅️ Back", key="btn_back_teritory", use_container_width=True, on_click=go_back_history)

    selected_month = f"{selected_year}-{selected_month_num}"
    target_excel_file = f"MASTER_{selected_month}.xlsx"
    if not os.path.exists(target_excel_file):
        matching_masters = glob.glob(f"MASTER_{selected_year}-*.xlsx")
        if matching_masters:
            matching_masters.sort(reverse=True)
            target_excel_file = matching_masters[0]
            selected_month = target_excel_file.replace("MASTER_", "").replace(".xlsx", "")
        elif os.path.exists("MASTER.xlsx"):
            target_excel_file = "MASTER.xlsx"

df_frx_raw = pd.DataFrame()
df_site = pd.DataFrame()
df_visit = pd.DataFrame()
df_atm_summary = pd.DataFrame()
daily_sheets = {}

if menu_option != "Jadwal Standby CSE":
    if os.path.exists(target_excel_file):
        (
            df_frx_raw,
            df_site,
            df_visit,
            df_atm_summary,
            daily_sheets,
        ) = load_data(target_excel_file)

def get_3_months_visit(current_month_str, wsid):
    parts = current_month_str.split("-")
    y, m = int(parts[0]), int(parts[1])

    dfs_to_concat = []
    for i in range(3):
        m_calc = m - i
        y_calc = y
        if m_calc <= 0:
            m_calc += 12
            y_calc -= 1

        m_str_file = f"MASTER_{y_calc}-{m_calc:02d}.xlsx"
        if os.path.exists(m_str_file):
            try:
                xls_m = pd.ExcelFile(m_str_file)
                target_sheet_visit = None
                for s in xls_m.sheet_names:
                    if "mainvisit" in s.lower():
                        target_sheet_visit = s
                        break
                
                if target_sheet_visit:
                    df_v = pd.read_excel(xls_m, target_sheet_visit)
                    if not df_v.empty:
                        df_v.columns = [str(c).strip().upper().replace(" ", "_") for c in df_v.columns]
                        col_id = None
                        for c_name in ["ID", "WSID", "WS_ID"]:
                            if c_name in df_v.columns:
                                col_id = c_name
                                break
                        if col_id:
                            df_f = df_v[
                                df_v[col_id].astype(str).str.upper()
                                == str(wsid).upper()
                            ]
                            if not df_f.empty:
                                dfs_to_concat.append(df_f)
            except Exception:
                pass

    if dfs_to_concat:
        combined = pd.concat(dfs_to_concat, ignore_index=True)
        combined = combined.drop_duplicates()
        return combined
    return pd.DataFrame()


EXCLUDED_CSE_KEYWORDS = [
    "PROVINSI", "PRIVINSI", "BALI", "NUSA TENGGARA BARAT", "NUSA TENGGARA TIMUR",
    "SERVICE AREA", "BIMA", "DENPASAR", "ENDE", "KUPANG", "MATARAM", "NEGARA",
    "LOMBOK TIMUR", "SINGARAJA", "SUMBAWA", "JATIM 1", "JATIM 2", "SULAMPUA 1", "SULAMPUA 2",
]

list_cse = ["-- Semua CSE --"]
if menu_option in ["Uptime CSE by Tipe Mesin", "Uptime Harian"]:
    try:
        db_file_conn = target_excel_file.replace(".xlsx", "_db.db")
        if os.path.exists(db_file_conn):
            conn = sqlite3.connect(db_file_conn)
            tables = [row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
            t_sheet = None
            for t in tables:
                if "bali_nusra" in t or "uptime_harian" in t or "uptime_cse" in t or ("uptime" in t and ("harian" in t or "daily" in t or "crm" in t)):
                    t_sheet = t
                    break
            if t_sheet:
                df_temp_cse = pd.read_sql(f"SELECT * FROM {t_sheet}", conn)
                for _, r in df_temp_cse.iterrows():
                    r_vals = [str(x).strip() for x in r.values if pd.notnull(x)]
                    row_str = " ".join([v.upper() for v in r_vals])
                    if len(r_vals) > 1 and "TOTAL" not in row_str and "UPTIME" not in row_str and "ENGINEER" not in row_str and "PENGELOLA" not in row_str:
                        col_no_val = str(r.iloc[0]).strip()
                        cse_name = str(r.iloc[1]).strip()
                        cse_upper = cse_name.upper()
                        
                        is_valid_row = False
                        try:
                            if float(col_no_val) > 0:
                                is_valid_row = True
                        except Exception:
                            pass

                        is_excluded = any(exc in cse_upper for exc in EXCLUDED_CSE_KEYWORDS)
                        if is_valid_row and cse_name and cse_name not in ["nan", "None", "-"] and not is_excluded and cse_name not in list_cse:
                            list_cse.append(cse_name)
            conn.close()
    except Exception:
        pass

list_pkt = ["-- Semua PKT --"]
if menu_option == "Uptime PKT by Tipe Mesin":
    try:
        db_file_conn = target_excel_file.replace(".xlsx", "_db.db")
        if os.path.exists(db_file_conn):
            conn = sqlite3.connect(db_file_conn)
            tables = [row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
            sheet_4_name = "pkt_tipe_mesin"
            if sheet_4_name not in tables:
                for t in tables:
                    if "pkt" in t:
                        sheet_4_name = t
                        break
            if sheet_4_name in tables:
                df_temp_pkt = pd.read_sql(f"SELECT * FROM {sheet_4_name}", conn)
                for _, r in df_temp_pkt.iterrows():
                    r_vals = [str(x).strip() for x in r.values if pd.notnull(x)]
                    row_str = " ".join([v.upper() for v in r_vals])
                    if not r_vals or "UPTIME" in row_str or "PROVINSI" in row_str or "SERVICE AREA" in row_str:
                        continue
                    if len(r_vals) > 1 and "TOTAL" not in row_str:
                        col_no_val = str(r.iloc[0]).strip()
                        pkt_name = str(r.iloc[1]).strip()
                        is_valid_no = False
                        try:
                            if float(col_no_val) > 0:
                                is_valid_no = True
                        except Exception:
                            pass
                        if is_valid_no and pkt_name and pkt_name not in ["nan", "None", "-"] and pkt_name not in list_pkt:
                            list_pkt.append(pkt_name)
            conn.close()
    except Exception:
        pass

list_wsid = []
if menu_option not in ["Teritori Mesin Engineer", "Jadwal Standby CSE", "Uptime PKT by Tipe Mesin"]:
    if not df_site.empty and "ID" in df_site.columns:
        list_wsid = df_site["ID"].dropna().unique().tolist()
    elif not df_atm_summary.empty and "WSID" in df_atm_summary.columns:
        list_wsid = df_atm_summary["WSID"].dropna().unique().tolist()

    if "ZTR4" not in list_wsid and list_wsid:
        list_wsid.insert(0, "ZTR4")

selected_wsid = "ZTR4"
selected_cse = "-- Semua CSE --"
selected_pkt = "-- Semua PKT --"
selected_daily_cse = "-- Semua CSE --"

if menu_option not in ["Teritori Mesin Engineer", "Jadwal Standby CSE"]:
    with col_wsid:
        if menu_option == "Uptime CSE by Tipe Mesin":
            if list_cse:
                selected_cse = st.selectbox(
                    "Pilih / Ketik CSE:",
                    options=list_cse,
                    index=0,
                    key="main_cse_select",
                )
        elif menu_option == "Uptime PKT by Tipe Mesin":
            if list_pkt:
                selected_pkt = st.selectbox(
                    "Pilih / Ketik PKT:",
                    options=list_pkt,
                    index=0,
                    key="main_pkt_select",
                )
        elif menu_option == "Uptime Harian":
            if list_cse:
                selected_daily_cse = st.selectbox(
                    "Pilih CSE / Pengelola:",
                    options=list_cse,
                    index=0,
                    key="main_daily_cse_select",
                )
        else:
            if list_wsid:
                default_idx = 0
                if "selected_wsid_from_teritory" in st.session_state:
                    if st.session_state["selected_wsid_from_teritory"] in list_wsid:
                        default_idx = list_wsid.index(st.session_state["selected_wsid_from_teritory"])
                
                selected_wsid = st.selectbox(
                    "Pilih / Ketik WSID:",
                    options=list_wsid,
                    index=default_idx,
                    key="main_wsid_select",
                )
                if "selected_wsid_from_teritory" in st.session_state:
                    del st.session_state["selected_wsid_from_teritory"]

    with col_btn:
        def go_to_visit():
            change_page("Riwayat Kunjungan (Visit)")

        if menu_option not in ["Uptime Harian", "Uptime CSE by Tipe Mesin", "Uptime PKT by Tipe Mesin"]:
            st.write("") # Memberi spasi vertikal agar sejajar dengan label selectbox di kolom sebelah
            st.button(
                f"🔗 Riwayat Kunjungan ({selected_wsid})",
                type="primary",
                key="btn_shortcut_visit",
                on_click=go_to_visit,
            )

# =============================================================
# FUNGSI MASTER CACHE HALAMAN 1 (STATUS TRX DARI ATM TOTAL)
# =============================================================
@st.cache_data(show_spinner=False)
def get_halaman1_cached_data(target_excel_file, selected_wsid, selected_month):
    info_data = {
        "Wsid": selected_wsid,
        "Ws Name": "-", "Trx": "-", "Serial No": "-", "MA_Company": "-",
        "Address": "-", "City": "-", "Province": "-", "Island": "-",
        "Location": "-", "SVC_Area": "-", "Mtype": "CRM", "Model": "-",
        "Sw Installed": "-", "Vendor": "-", "Engineer": "-",
        "SE Name": "-", "SE Code": "-", "Bank Code": "-", "Bank Name": "-",
        "Ops Hour": "-", "Ops Hour Range": "-", "Svc Hour Range": "-",
        "Ma Status": "-", "Zone": "-", "Production Year": "-",
        "Installation Date": "-", "Warranty Exp date": "-", "Gps Address": "-",
        "Gps Lat": "-", "Gps Long": "-"
    }
    
    freq_dt_val = 0
    achieve_ut_float = 0.0
    chart_dates = []
    chart_uptime = []
    table_frx_rows_html = []
    daily_frx_list = []
    status_trx_val = "-"
    year_str, month_str = selected_month.split("-")

    db_file_conn = target_excel_file.replace(".xlsx", "_db.db")
    if os.path.exists(db_file_conn):
        conn = sqlite3.connect(db_file_conn)
        all_tables = [row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
        
        target_site_tbl = next((t for t in all_tables if "site" in t.lower() or "blns" in t.lower()), None)
        
        if target_site_tbl:
            df_site_sql = pd.read_sql(f"SELECT * FROM {target_site_tbl}", conn)
            if not df_site_sql.empty:
                df_site_sql.columns = [str(c).strip().upper().replace(" ", "_").replace(".", "_") for c in df_site_sql.columns]
                
                for _, r_b in df_site_sql.iterrows():
                    w_val_0 = str(r_b.iloc[0]).strip().upper() if len(r_b) > 0 and pd.notnull(r_b.iloc[0]) else ""
                    w_val_1 = str(r_b.iloc[1]).strip().upper() if len(r_b) > 1 and pd.notnull(r_b.iloc[1]) else ""
                    
                    target_ws = str(selected_wsid).strip().upper()
                    
                    if w_val_0 == target_ws or w_val_1 == target_ws or target_ws in w_val_0 or target_ws in w_val_1:
                        info_data["Ws Name"] = str(r_b.iloc[2]).strip() if len(r_b) > 2 and pd.notnull(r_b.iloc[2]) else "-"
                        info_data["Serial No"] = str(r_b.iloc[3]).strip() if len(r_b) > 3 and pd.notnull(r_b.iloc[3]) else "-"
                        info_data["Address"] = str(r_b.iloc[4]).strip() if len(r_b) > 4 and pd.notnull(r_b.iloc[4]) else "-"
                        info_data["City"] = str(r_b.iloc[5]).strip() if len(r_b) > 5 and pd.notnull(r_b.iloc[5]) else "-"
                        info_data["Province"] = str(r_b.iloc[6]).strip() if len(r_b) > 6 and pd.notnull(r_b.iloc[6]) else "-"
                        info_data["Island"] = str(r_b.iloc[7]).strip() if len(r_b) > 7 and pd.notnull(r_b.iloc[7]) else "-"
                        info_data["Location"] = str(r_b.iloc[8]).strip() if len(r_b) > 8 and pd.notnull(r_b.iloc[8]) else "-"
                        info_data["Mtype"] = str(r_b.iloc[10]).strip().upper() if len(r_b) > 10 and pd.notnull(r_b.iloc[10]) else "CRM"
                        info_data["Model"] = str(r_b.iloc[11]).strip() if len(r_b) > 11 and pd.notnull(r_b.iloc[11]) else "-"
                        info_data["Vendor"] = str(r_b.iloc[13]).strip() if len(r_b) > 13 and pd.notnull(r_b.iloc[13]) else "-"
                        info_data["Bank Code"] = str(r_b.iloc[14]).strip() if len(r_b) > 14 and pd.notnull(r_b.iloc[14]) else "-"
                        info_data["Bank Name"] = str(r_b.iloc[15]).strip() if len(r_b) > 15 and pd.notnull(r_b.iloc[15]) else "-"
                        info_data["MA_Company"] = str(r_b.iloc[16]).strip() if len(r_b) > 16 and pd.notnull(r_b.iloc[16]) else "-"
                        
                        ops_col = next((c for c in df_site_sql.columns if "OPS_HOUR" in c or "OPR_HOUR" in c or c == "R"), None)
                        if ops_col and pd.notnull(r_b[ops_col]):
                            val_ops = str(r_b[ops_col]).strip()
                            info_data["Ops Hour"] = val_ops if val_ops.lower() not in ["nan", "none", "", "-"] else "-"
                        else:
                            info_data["Ops Hour"] = str(r_b.iloc[17]).strip() if len(r_b) > 17 and pd.notnull(r_b.iloc[17]) and str(r_b.iloc[17]).strip().lower() not in ["nan", "none", ""] else "-"
                        
                        info_data["Ops Hour Range"] = str(r_b.iloc[18]).strip() if len(r_b) > 18 and pd.notnull(r_b.iloc[18]) else "-"
                        
                        svc_col = next((c for c in df_site_sql.columns if "SVC" in c or "SERVICE" in c or c == "T"), None)
                        val_svc = "-"
                        if svc_col and pd.notnull(r_b[svc_col]):
                            v = str(r_b[svc_col]).strip()
                            if v.lower() not in ["nan", "none", "", "-"]:
                                val_svc = v
                        if val_svc == "-" and len(r_b) > 19 and pd.notnull(r_b.iloc[19]):
                            v = str(r_b.iloc[19]).strip()
                            if v.lower() not in ["nan", "none", "", "-"]:
                                val_svc = v
                        if val_svc == "-":
                            for idx_check in [19, 20, 21]:
                                if len(r_b) > idx_check and pd.notnull(r_b.iloc[idx_check]):
                                    v = str(r_b.iloc[idx_check]).strip()
                                    if ":" in v and v.lower() not in ["nan", "none", "", "-"]:
                                        val_svc = v
                                        break
                        info_data["Svc Hour Range"] = val_svc
                        
                        info_data["Ma Status"] = str(r_b.iloc[21]).strip() if len(r_b) > 21 and pd.notnull(r_b.iloc[21]) else "-"
                        info_data["SE Code"] = str(r_b.iloc[22]).strip() if len(r_b) > 22 and pd.notnull(r_b.iloc[22]) else "-"
                        info_data["SE Name"] = str(r_b.iloc[23]).strip() if len(r_b) > 23 and pd.notnull(r_b.iloc[23]) else "-"
                        info_data["SVC_Area"] = str(r_b.iloc[26]).strip() if len(r_b) > 26 and pd.notnull(r_b.iloc[26]) else "-"
                        info_data["Zone"] = str(r_b.iloc[27]).strip() if len(r_b) > 27 and pd.notnull(r_b.iloc[27]) else "-"
                        info_data["Production Year"] = str(r_b.iloc[30]).strip() if len(r_b) > 30 and pd.notnull(r_b.iloc[30]) else "-"
                        info_data["Installation Date"] = str(r_b.iloc[31]).strip() if len(r_b) > 31 and pd.notnull(r_b.iloc[31]) else "-"
                        info_data["Sw Installed"] = str(r_b.iloc[32]).strip() if len(r_b) > 32 and pd.notnull(r_b.iloc[32]) else "-"
                        info_data["Warranty Exp date"] = str(r_b.iloc[33]).strip() if len(r_b) > 33 and pd.notnull(r_b.iloc[33]) else "-"
                        
                        info_data["Gps Lat"] = str(r_b.iloc[39]).strip() if len(r_b) > 39 and pd.notnull(r_b.iloc[39]) else "-"
                        info_data["Gps Long"] = str(r_b.iloc[40]).strip() if len(r_b) > 40 and pd.notnull(r_b.iloc[40]) else "-"
                        
                        gps_col = next((c for c in df_site_sql.columns if "GPS" in c or "ADDRESS" in c or "ALAMAT" in c), None)
                        val_gps = "-"
                        if gps_col and pd.notnull(r_b[gps_col]):
                            v = str(r_b[gps_col]).strip()
                            if v.lower() not in ["nan", "none", "", "-"]:
                                val_gps = v
                        if val_gps == "-":
                            for idx_check in [43, 44, 45, 42]:
                                if len(r_b) > idx_check and pd.notnull(r_b.iloc[idx_check]):
                                    v = str(r_b.iloc[idx_check]).strip()
                                    if len(v) > 5 and v.lower() not in ["nan", "none", "", "-"]:
                                        val_gps = v
                                        break
                        info_data["Gps Address"] = val_gps
                        info_data["Engineer"] = info_data["SE Name"]
                        break

        target_atm_t = next((t for t in all_tables if "atm" in t.lower() and "total" in t.lower()), None)
        if target_atm_t:
            df_atm_summary = pd.read_sql(f"SELECT * FROM [{target_atm_t}]", conn)
            if not df_atm_summary.empty:
                df_atm_summary.columns = [str(c).strip().upper().replace(" ", "_").replace(".", "_") for c in df_atm_summary.columns]
                
                ut_col_name = next((c for c in df_atm_summary.columns if c in ["UPTIME", "ACHIEVEMENT", "%", "ACHIEV"] or "UPTIME" in c), None)
                frx_col_name = next((c for c in df_atm_summary.columns if c in ["FRX", "FREQUENCY", "FREQ", "TOTAL_FRX"] or ("FRX" in c and "TOTAL" in c)), None)
                col_w_sum = next((c for c in ["WSID", "ID", "WS_ID", "ATM_ID"] if c in df_atm_summary.columns), None)

                if col_w_sum:
                    match_sum = df_atm_summary[df_atm_summary[col_w_sum].astype(str).str.upper().str.strip() == str(selected_wsid).strip().upper()]
                    if not match_sum.empty:
                        r_sum = match_sum.iloc[0]
                        if ut_col_name and pd.notnull(r_sum[ut_col_name]):
                            try:
                                val_u = float(str(r_sum[ut_col_name]).strip())
                                achieve_ut_float = val_u * 100 if val_u <= 1.0 else val_u
                            except: pass
                        if frx_col_name and pd.notnull(r_sum[frx_col_name]):
                            try: freq_dt_val = int(float(str(r_sum[frx_col_name]).strip()))
                            except: pass

            df_atm_raw = pd.read_sql(f"SELECT * FROM [{target_atm_t}]", conn)
            if not df_atm_raw.empty:
                target_w = str(selected_wsid).strip().upper()
                for _, r in df_atm_raw.iterrows():
                    v0 = str(r.iloc[0]).strip().upper() if len(r) > 0 and pd.notnull(r.iloc[0]) else ""
                    v1 = str(r.iloc[1]).strip().upper() if len(r) > 1 and pd.notnull(r.iloc[1]) else ""
                    if target_w == v0 or target_w == v1 or target_w in v0 or target_w in v1:
                        if len(r) > 39 and pd.notnull(r.iloc[39]):
                            raw_val = str(r.iloc[39]).strip()
                            if raw_val != "" and raw_val.lower() not in ["nan", "none", "-", "0"]:
                                status_trx_val = raw_val
                        break

        sorted_days = sorted([int(t.replace("day_", "")) for t in all_tables if t.startswith("day_")])
        for d in sorted_days:
            df_d = pd.read_sql(f"SELECT * FROM day_{d}", conn)
            if df_d.empty: continue
            df_d.columns = [str(c).strip().upper().replace(" ", "_") for c in df_d.columns]
            col_w_day = "WSID" if "WSID" in df_d.columns else ("ID" if "ID" in df_d.columns else df_d.columns[0])
            row_match = df_d[df_d[col_w_day].astype(str).str.upper() == str(selected_wsid).upper()]
            
            tgl_str = f"{d:02d}/{month_str}/{year_str}"
            chart_tgl = f"{d:02d}-{month_str}"
            
            if not row_match.empty:
                r = row_match.iloc[0]
                def get_val(key_list):
                    for k in key_list:
                        if k in df_d.columns and pd.notnull(r.get(k)):
                            try: return float(r.get(k))
                            except: pass
                    return 0.0

                hw_freq = int(get_val(["FRX", "FREQUENCY"]))
                hw_dur = int(get_val(["DUR", "DURATION"]))
                hw_dt = round(get_val(["%", "DT", "HW_%"]), 2)
                
                if hw_freq > 0: daily_frx_list.append(hw_freq)

                p01_freq, p01_dur, p01_dt = int(get_val(["FRX_1", "FRX.1"])), int(get_val(["DUR_1", "DUR.1"])), round(get_val(["%_1", "%.1"]), 2)
                p02_freq, p02_dur, p02_dt = int(get_val(["FRX_2", "FRX.2"])), int(get_val(["DUR_2", "DUR.2"])), round(get_val(["%_2", "%.2"]), 2)
                p03_freq, p03_dur, p03_dt = int(get_val(["FRX_3", "FRX.3"])), int(get_val(["DUR_3", "DUR.3"])), round(get_val(["%_3", "%.3"]), 2)
                p04_freq, p04_dur, p04_dt = int(get_val(["FRX_4", "FRX.4"])), int(get_val(["DUR_4", "DUR.4"])), round(get_val(["%_4", "%.4"]), 2)

                uptime_val = round(100 - hw_dt, 2) if hw_dt <= 100 else 0.0
                chart_dates.append(chart_tgl)
                chart_uptime.append(uptime_val)

                row_bg = "#fcdad7" if (hw_freq > 0 or hw_dur > 0) else "#d5f5e3"
                dt_color = "#c0392b" if hw_dt > 0 else "#27ae60"

                row_html = f"""
                <tr style="background-color: {row_bg}; color: #000000;">
                    <td style="font-weight:bold; text-align:center;">{tgl_str}</td>
                    <td>{hw_freq}</td><td>{hw_dur}</td><td style="color:{dt_color}; font-weight:bold;">{hw_dt}</td>
                    <td>{p01_freq}</td><td>{p01_dur}</td><td>{p01_dt}</td>
                    <td>{p02_freq}</td><td>{p02_dur}</td><td>{p02_dt}</td>
                    <td>{p03_freq}</td><td>{p03_dur}</td><td>{p03_dt}</td>
                    <td>{p04_freq}</td><td>{p04_dur}</td><td>{p04_dt}</td>
                </tr>
                """
                table_frx_rows_html.append(row_html)
            else:
                table_frx_rows_html.append(f'<tr style="background-color: #ffffff;"><td style="font-weight:bold; text-align:center;">{tgl_str}</td><td colspan="15"></td></tr>')
        conn.close()

    if freq_dt_val == 0 and daily_frx_list:
        freq_dt_val = sum(daily_frx_list)
    if achieve_ut_float == 0.0 and chart_uptime:
        achieve_ut_float = sum(chart_uptime) / len(chart_uptime)

    return info_data, freq_dt_val, achieve_ut_float, chart_dates, chart_uptime, "".join(table_frx_rows_html), status_trx_val

# =============================================================
# HALAMAN 1: UPTIME FRX DT (TAMPILAN UTAMA)
# =============================================================
if menu_option == "Uptime FRX DT (WSID)":
    
    info_data, freq_dt_val, achieve_ut_float, chart_dates, chart_uptime, table_frx_body, status_trx_val = get_halaman1_cached_data(target_excel_file, selected_wsid, selected_month)
    
    achieve_ut_str = f"{round(achieve_ut_float, 2)}%"
    mtype = info_data.get("Mtype", "CRM")
    target_ut = 99.75 if "ATM" in str(mtype).upper() else 99.20
    is_tercapai = achieve_ut_float >= target_ut

    status_text = "Tercapai" if is_tercapai else "Tidak tercapai"
    status_color = "#27ae60" if is_tercapai else "#e74c3c"
    ut_color = status_color
    frx_color = "#e74c3c" if freq_dt_val > 0 else "#27ae60"

    trx_lower = status_trx_val.lower()
    if "sangat tinggi" in trx_lower:
        badge_color = "#e74c3c"
    elif "tinggi" in trx_lower:
        badge_color = "#e67e22"
    elif "average" in trx_lower:
        badge_color = "#f39c12"
    elif "cukup rendah" in trx_lower:
        badge_color = "#ff69b4"
    elif "rendah" in trx_lower:
        badge_color = "#1b4f72"
    else:
        badge_color = "#7f8c8d"

    df_3m_visit = get_3_months_visit(selected_month, selected_wsid) if 'get_3_months_visit' in globals() else pd.DataFrame()

    history_rows = []
    if not df_3m_visit.empty:
        df_3m_visit.columns = [str(c).strip().upper().replace(" ", "_") for c in df_3m_visit.columns]
        if "STARTED" in df_3m_visit.columns:
            df_3m_visit = df_3m_visit.sort_values(by="STARTED", ascending=True)

        for _, vrow in df_3m_visit.iterrows():
            spart_val = vrow.get("SPART", None)
            if pd.notnull(spart_val) and str(spart_val).strip() not in ["", "0", "nan", "None"]:
                start_time = vrow.get("STARTED", "-")
                end_time = vrow.get("FINISHED", "-")
                history_rows.append({"start": str(start_time), "end": str(end_time), "spart": str(spart_val).strip()})

    history_html = (
        "".join([f"<tr><td>{h['start']}</td><td>{h['end']}</td><td>{h['spart']}</td></tr>" for h in history_rows])
        if history_rows
        else '<tr><td colspan="3" style="text-align:center; color:#888; padding:15px;">Tidak ada history penggantian part dalam 3 bulan terakhir</td></tr>'
    )

    lat_val = info_data.get('Gps Lat', '-7.250445')
    lon_val = info_data.get('Gps Long', '112.768845')
    gps_addr = info_data.get('Gps Address', '-')
    
    try: lat_float = float(lat_val)
    except: lat_float = -8.409518
        
    try: lon_float = float(lon_val)
    except: lon_float = 115.188919

    osm_link = f"https://www.openstreetmap.org/?mlat={lat_float}&mlon={lon_float}#map=17/{lat_float}/{lon_float}"
    gmaps_link = f"https://www.google.com/maps/search/?api=1&query={lat_float},{lon_float}"

    html_code = f"""
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
        <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
        <style>
            * {{ box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; }}
            body {{ background-color: #f4f6f9; padding: 10px; color: #333; }}
            .top-banner {{ background-color: #ffffff; border-radius: 6px; padding: 12px 20px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 2px 4px rgba(0,0,0,0.08); margin-bottom: 15px; font-weight: 600; font-size: 15px; }}
            .uptime-box {{ border: 2px solid #e74c3c; background-color: #fdfefe; padding: 8px 18px; border-radius: 6px; font-weight: 800; font-size: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
            .dashboard-grid {{ display: grid; grid-template-columns: 28% 44% 28%; gap: 15px; }}
            .card {{ background: #ffffff; border-radius: 6px; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.08); margin-bottom: 15px; }}
            .card-header {{ padding: 10px 15px; color: #ffffff; font-weight: bold; display: flex; justify-content: space-between; align-items: center; font-size: 14px; cursor: pointer; user-select: none; }}
            .card-header.lightblue {{ background-color: #3598db; }}
            .card-header.green {{ background-color: #27ae60; }}
            .card-header.purple {{ background-color: #9b59b6; }}
            .card-header.gold {{ background-color: #f1c40f; color: #000000; }}
            .right-column-grid {{ display: flex; flex-direction: column; gap: 15px; }}
            .history-card-container {{ height: 680px; display: flex; flex-direction: column; margin-bottom: 0 !important; }}
            .history-scroll-box {{ max-height: 610px; overflow-y: auto; }}
            .map-card-container {{ background: #ffffff; border-radius: 6px; box-shadow: 0 2px 5px rgba(0,0,0,0.08); overflow: hidden; height: 320px; margin-top: 15px; display: flex; flex-direction: column; }}
            .map-header-darkgreen {{ background-color: #1e8449; color: #ffffff; padding: 10px 15px; font-weight: bold; display: flex; justify-content: space-between; align-items: center; font-size: 14px; cursor: pointer; user-select: none; }}
            .map-content-area {{ padding: 0; flex-grow: 1; position: relative; display: flex; }}
            #osmMap {{ width: 55%; height: 100%; }}
            .map-info-panel {{ width: 45%; padding: 12px 15px; font-size: 11px; display: flex; flex-direction: column; justify-content: flex-start; background: #fafafa; border-left: 1px solid #eef2f5; overflow-y: auto; }}
            .map-info-row {{ margin-bottom: 10px; }}
            .map-info-label {{ font-weight: bold; color: #555; margin-bottom: 2px; text-transform: uppercase; font-size: 10px; }}
            .map-info-val {{ color: #222; line-height: 1.3; font-weight: 500; }}
            .map-links a {{ color: #3498db; text-decoration: none; font-weight: bold; margin-right: 8px; }}
            .header-engineer-info {{ color: #ffff00; font-size: 11px; font-weight: bold; margin-left: 8px; text-shadow: 0px 0px 2px rgba(0, 0, 0, 0.5); }}
            .info-table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
            .info-table td {{ padding: 7px 12px; border-bottom: 1px solid #eef2f5; }}
            .info-table tr td:first-child {{ font-weight: bold; color: #444; width: 35%; }}
            .history-table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
            .history-table th {{ background: #f8f9fa; text-align: left; padding: 8px; border-bottom: 2px solid #dee2e6; position: sticky; top: 0; z-index: 1; }}
            .history-table td {{ padding: 8px; border-bottom: 1px solid #eee; }}
            .frx-excel-table {{ width: 100%; border-collapse: collapse; font-size: 10px; text-align: center; }}
            .frx-excel-table th, .frx-excel-table td {{ border: 1px solid #e2b667; padding: 4px 2px; }}
            .p02-header-bg {{ background-color: #f39c12 !important; color: #000000 !important; font-size: 11px; font-weight: bold; }}
            .p02-sub-header {{ background-color: #fdf3e7 !important; color: #000000 !important; font-weight: bold; }}
            .table-scroll {{ max-height: none; overflow-y: visible; overflow-x: auto; position: relative; }}
        </style>
    </head>
    <body>
        <div class="top-banner">
            <div>WSID: <strong style="color: #2c3e50; font-size:17px;">{info_data['Wsid']}</strong></div>
            <div class="uptime-box">
                UPTIME BULANAN: {mtype}: <span style="color:{ut_color}; font-weight:800; font-size:16px;">{achieve_ut_str}</span> 
                &nbsp;|&nbsp; STATUS: <span style="color:{status_color}; font-weight:800;">{status_text}</span> 
                &nbsp;|&nbsp; TOTAL FRX: <strong style="color:{frx_color}; font-size:17px;">{freq_dt_val}</strong>
                &nbsp;|&nbsp; STATUS TRX: <span style="color: {badge_color}; font-weight: 800; font-size: 15px;">{status_trx_val}</span>
            </div>
            <div><span style="background-color:#eef2f7; padding:4px 10px; border-radius:4px; font-weight:bold; color:#2c3e50;">📅 {selected_month}</span></div>
        </div>

        <div class="dashboard-grid">
            <div class="card">
                <div class="card-header lightblue" onclick="toggleCard('infoContent', 'arrow1')"><span>Informasi Mesin</span><span id="arrow1">▼</span></div>
                <div id="infoContent"><table class="info-table">
                    <tr><td>Wsid</td><td>{info_data.get('Wsid', '-')}</td></tr>
                    <tr><td>Ws Name</td><td>{info_data.get('Ws Name', '-')}</td></tr>
                    <tr><td>Serial No</td><td>{info_data.get('Serial No', '-')}</td></tr>
                    <tr><td>MA_Company</td><td>{info_data.get('MA_Company', '-')}</td></tr>
                    <tr><td>SE Name</td><td>{info_data.get('SE Name', '-')}</td></tr>
                    <tr><td>SE Code</td><td>{info_data.get('SE Code', '-')}</td></tr>
                    <tr><td>Bank Code</td><td>{info_data.get('Bank Code', '-')}</td></tr>
                    <tr><td>Bank Name</td><td>{info_data.get('Bank Name', '-')}</td></tr>
                    <tr><td>Address</td><td>{info_data.get('Address', '-')}</td></tr>
                    <tr><td>City</td><td>{info_data.get('City', '-')}</td></tr>
                    <tr><td>Province</td><td>{info_data.get('Province', '-')}</td></tr>
                    <tr><td>Island</td><td>{info_data.get('Island', '-')}</td></tr>
                    <tr><td>Location</td><td>{info_data.get('Location', '-')}</td></tr>
                    <tr><td>SVC_Area</td><td>{info_data.get('SVC_Area', '-')}</td></tr>
                    <tr><td>Mtype</td><td>{info_data.get('Mtype', '-')}</td></tr>
                    <tr><td>Model</td><td>{info_data.get('Model', '-')}</td></tr>
                    <tr><td>Ops Hour</td><td>{info_data.get('Ops Hour', '-')}</td></tr>
                    <tr><td>Ops Hour Range</td><td>{info_data.get('Ops Hour Range', '-')}</td></tr>
                    <tr><td>Svc Hour Range</td><td>{info_data.get('Svc Hour Range', '-')}</td></tr>
                    <tr><td>Ma Status</td><td>{info_data.get('Ma Status', '-')}</td></tr>
                    <tr><td>Zone</td><td>{info_data.get('Zone', '-')}</td></tr>
                    <tr><td>Production Year</td><td>{info_data.get('Production Year', '-')}</td></tr>
                    <tr><td>Installation Date</td><td>{info_data.get('Installation Date', '-')}</td></tr>
                    <tr><td>Warranty Exp date</td><td>{info_data.get('Warranty Exp date', '-')}</td></tr>
                    <tr><td>Sw Installed</td><td>{info_data.get('Sw Installed', '-')}</td></tr>
                    <tr><td>Vendor</td><td>{info_data.get('Vendor', '-')}</td></tr>
                    <tr><td>Gps Address</td><td>{info_data.get('Gps Address', '-')}</td></tr>
                    <tr><td>Gps Lat</td><td>{info_data.get('Gps Lat', '-')}</td></tr>
                    <tr><td>Gps Long</td><td>{info_data.get('Gps Long', '-')}</td></tr>
                </table></div>
            </div>

            <div>
                <div class="card">
                    <div class="card-header green" onclick="toggleCard('chartContent', 'arrow2')"><span>Grafik Uptime Harian</span><span id="arrow2">▼</span></div>
                    <div id="chartContent" style="padding:10px;"><canvas id="lineChart" height="100"></canvas></div>
                </div>

                <div class="card">
                    <div class="card-header purple" onclick="toggleCard('tableContent', 'arrow3')">
                        <div><span>Uptime Harian</span><span class="header-engineer-info">| {info_data['Wsid']} | {info_data['Engineer']}</span></div>
                        <span id="arrow3">▼</span>
                    </div>
                    <div id="tableContent" class="table-scroll">
                        <table class="frx-excel-table">
                            <thead>
                                <tr class="p02-header-bg">
                                    <th rowspan="2">Tgl</th><th colspan="3">HW TOTAL</th><th colspan="3">P01 - Pick Modul</th><th colspan="3">P02 - Presenter</th><th colspan="3">P03 - Reject</th><th colspan="3">P04 - Card Reader</th>
                                </tr>
                                <tr class="p02-sub-header">
                                    <th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th>
                                </tr>
                            </thead>
                            <tbody>{table_frx_body}</tbody>
                        </table>
                    </div>
                </div>
            </div>

            <div class="right-column-grid">
                <div class="card history-card-container">
                    <div class="card-header gold" onclick="toggleCard('historyContent', 'arrow4')"><span>History Spart (3 Bulan Terakhir)</span><span id="arrow4">▼</span></div>
                    <div id="historyContent" class="history-scroll-box">
                        <table class="history-table">
                            <thead><tr><th>Start Time</th><th>End Time</th><th>SPart</th></tr></thead>
                            <tbody>{history_html}</tbody>
                        </table>
                    </div>
                </div>

                <div class="map-card-container">
                    <div class="map-header-darkgreen" onclick="toggleCard('mapContent', 'arrowMap')">
                        <span>Lokasi Mesin (OpenStreetMap)</span><span id="arrowMap">▼</span>
                    </div>
                    <div id="mapContent" class="map-content-area">
                        <div id="osmMap"></div>
                        <div class="map-info-panel">
                            <div class="map-info-row">
                                <div class="map-info-label">Alamat</div>
                                <div class="map-info-val">{gps_addr}</div>
                            </div>
                            <div class="map-info-row">
                                <div class="map-info-label">Koordinat</div>
                                <div class="map-info-val">{lat_val}, {lon_val}</div>
                            </div>
                            <div class="map-info-row">
                                <div class="map-info-label">Buka di Maps</div>
                                <div class="map-info-val map-links">
                                    <a href="{osm_link}" target="_blank">OpenStreetMap</a> &bull; 
                                    <a href="{gmaps_link}" target="_blank">Google Maps</a>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <script>
            function toggleCard(cId, aId) {{
                const c = document.getElementById(cId); const a = document.getElementById(aId);
                if (c.style.display === "none") {{ c.style.display = "block"; a.innerText = "▼"; }} else {{ c.style.display = "none"; a.innerText = "▲"; }}
            }}

            new Chart(document.getElementById('lineChart'), {{
                type: 'line',
                data: {{ labels: {json.dumps(chart_dates)}, datasets: [{{ label: 'Uptime (%)', data: {json.dumps(chart_uptime)}, borderColor: '#27ae60', backgroundColor: 'rgba(39, 174, 96, 0.1)', borderWidth: 2, fill: true }}] }},
                options: {{ responsive: true, plugins: {{ legend: {{ display: false }} }}, scales: {{ y: {{ min: 70, max: 100 }} }} }}
            }});

            document.addEventListener("DOMContentLoaded", function() {{
                var lat = {lat_float};
                var lon = {lon_float};
                var map = L.map('osmMap', {{ zoomControl: true }}).setView([lat, lon], 16);
                
                L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
                    maxZoom: 19,
                    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>'
                }}).addTo(map);

                L.marker([lat, lon]).addTo(map)
                  .bindPopup("<b>{info_data.get('Wsid', '-')}</b><br>{gps_addr}")
                  .openPopup();
                
                setTimeout(function() {{ map.invalidateSize(); }}, 200);
            }});
        </script>
    </body>
    </html>
    """

    st.components.v1.html(html_code, height=1080, scrolling=True)

# =============================================================
# HALAMAN 2: UPTIME HARIAN
# =============================================================
if menu_option == "Uptime Harian":
    col_t, col_b = st.columns([5, 1])
    with col_t:
        filter_info_text = f" - Filter CSE / Pengelola: <span style='color: #3498db;'>{selected_daily_cse}</span>" if selected_daily_cse != "-- Semua CSE --" else ""
        st.markdown(f"<h4 style='margin-bottom: 12px; color: #2c3e50;'>Uptime Harian ({selected_month}){filter_info_text}</h4>", unsafe_allow_html=True)
    with col_b:
        st.write("")
        st.button("⬅️ Back", key="btn_back_daily", use_container_width=True, on_click=go_back_history)

    try:
        table_name_sql = f"uptime_harian_{selected_month.replace('-', '_')}"
        current_dir = os.path.dirname(os.path.abspath(__file__))
        db_path = None
        try:
            all_files = os.listdir(current_dir)
            master_dbs = [f for f in all_files if 'MASTER' in f and '_db' in f]
            if master_dbs:
                master_dbs.sort()
                db_path = os.path.join(current_dir, master_dbs[-1])
            elif 'atm_database' in all_files:
                db_path = os.path.join(current_dir, 'atm_database')
        except Exception:
            pass

        df_sheet2 = pd.DataFrame()
        if db_path and os.path.exists(db_path):
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name_sql}'")
                table_exists = cursor.fetchone()
                
                if table_exists:
                    df_sheet2 = pd.read_sql(f"SELECT * FROM {table_name_sql}", conn)
                conn.close()
            except Exception:
                pass

        if df_sheet2.empty:
            xls = pd.ExcelFile(target_excel_file)
            sheet_2_name = None
            for s in xls.sheet_names:
                if "UPTIME" in s.upper() and ("HARIAN" in s.upper() or "DAILY" in s.upper() or "CRM" in s.upper()):
                    sheet_2_name = s
                    break
            if not sheet_2_name and len(xls.sheet_names) > 1:
                sheet_2_name = xls.sheet_names[1]

            if sheet_2_name:
                df_sheet2 = pd.read_excel(xls, sheet_name=sheet_2_name, header=None)
                if db_path and os.path.exists(db_path):
                    try:
                        conn = sqlite3.connect(db_path)
                        df_sheet2.to_sql(table_name_sql, conn, if_exists='replace', index=False)
                        conn.close()
                    except Exception:
                        pass

        if not df_sheet2.empty:
            ALL_MONTHS = ["JAN", "FEB", "MAR", "APR", "MEI", "JUN", "JUL", "AGUS", "SEPT", "OKT", "NOV", "DES"]
            selected_m_num = int(selected_month.split("-")[1])
            start_idx = max(0, selected_m_num - 3)
            allowed_months = ALL_MONTHS[start_idx:selected_m_num]

            def build_uptime_daily_html(df, keep_months, filter_cse):
                html = """
                <style>
                    .excel-wrapper { display: flex; flex-direction: column; gap: 12px; }
                    .excel-container { overflow-x: auto; overflow-y: visible !important; max-height: none !important; border: 1px solid #7f8c8d; border-radius: 2px; background-color: #ffffff; }
                    .excel-table { width: 100%; border-collapse: collapse; font-size: 11px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
                    .excel-table td, .excel-table th { border: 1px solid #a6a6a6; padding: 4px 6px; white-space: nowrap; text-align: center; }
                    .header-row { background-color: #8ea9db !important; color: #000000 !important; font-weight: bold; position: sticky; top: 0; z-index: 2; }
                    .col-no { width: 32px !important; text-align: center; }
                    .col-cse { text-align: left !important; padding-left: 6px !important; font-weight: 500; }
                    .row-section-title { background-color: #e2efda !important; color: #000000 !important; font-weight: bold; text-align: left !important; padding: 6px 10px !important; font-size: 11px; }
                    .bg-green { background-color: #c6efce !important; color: #006100 !important; font-weight: bold; }
                    .bg-red { background-color: #ffc7ce !important; color: #9c0006 !important; font-weight: bold; }
                    .row-summary { background-color: #f2f2f2 !important; font-weight: bold; }
                </style>
                <div class="excel-wrapper">
                """

                valid_col_indices = []
                header_row_idx = None

                for r_idx, r in df.iterrows():
                    r_vals = ["" if pd.isnull(x) else str(x).strip().upper() for x in r.values]
                    if "NO" in r_vals and ("CSE" in r_vals or "PKT" in r_vals):
                        header_row_idx = r_idx
                        for col_i, col_v in enumerate(r_vals):
                            if col_i in [0, 1, 2]:
                                valid_col_indices.append(col_i)
                            elif col_v in ALL_MONTHS:
                                if col_v in keep_months:
                                    valid_col_indices.append(col_i)
                        break

                if header_row_idx is not None:
                    header_vals = ["" if pd.isnull(x) else str(x).strip().upper() for x in df.iloc[header_row_idx].values]
                    cse_start_row = header_row_idx + 1
                    cse_end_row = len(df)
                    for r_i in range(cse_start_row, len(df)):
                        r_txt = " ".join([str(x).upper() for x in df.iloc[r_i].values if pd.notnull(x)])
                        if "TOTAL" in r_txt or "PENGELOLA" in r_txt or "BALI NUSRA" in r_txt:
                            cse_end_row = r_i
                            break

                    for col_i in range(len(header_vals)):
                        if col_i in valid_col_indices:
                            continue
                        col_head_str = header_vals[col_i]
                        is_day_col = False
                        try:
                            f_v = float(col_head_str)
                            if 1 <= f_v <= 31:
                                is_day_col = True
                        except Exception:
                            pass

                        if is_day_col:
                            has_real_data = False
                            for r_idx in range(cse_start_row, cse_end_row):
                                cell_val = df.iloc[r_idx, col_i]
                                if pd.notnull(cell_val):
                                    s_val = str(cell_val).strip().upper()
                                    if s_val not in ["", "#DIV/0!", "NAN", "NONE"]:
                                        try:
                                            if float(s_val) > 0:
                                                has_real_data = True
                                                break
                                        except Exception:
                                            has_real_data = True
                                            break
                            if has_real_data:
                                valid_col_indices.append(col_i)

                rows_engineer = []
                rows_pengelola = []
                current_section = "ENGINEER"
                is_filtering = (filter_cse and filter_cse != "-- Semua CSE --")

                for row_idx, row in df.iterrows():
                    row_vals = ["" if pd.isnull(x) else str(x).strip() for x in row.values]
                    if not any(row_vals):
                        continue
                    row_str = " ".join([v.upper() for v in row_vals])

                    if row_idx == 0 and "UPTIME DAILY" in row_str:
                        continue

                    if "ENGINEER" in row_str or "UT :" in row_str:
                        current_section = "ENGINEER"
                        continue
                    if "PENGELOLA" in row_str:
                        current_section = "PENGELOLA"
                        continue

                    if is_filtering and current_section == "PENGELOLA":
                        continue

                    match_filter = True
                    if is_filtering:
                        is_header = ("NO" in [v.upper() for v in row_vals]) and ("CSE" in [v.upper() for v in row_vals] or "PKT" in [v.upper() for v in row_vals])
                        is_total_or_sec = any(k in row_str for k in ["TOTAL", "PENGELOLA", "ENGINEER", "ID TIDAK TERCAPAI"])
                        
                        if is_header:
                            match_filter = True
                        elif is_total_or_sec:
                            match_filter = False
                        else:
                            row_cse_name = str(row.iloc[1]).strip() if len(row_vals) > 1 else ""
                            if row_cse_name.upper() == filter_cse.upper():
                                match_filter = True
                            else:
                                match_filter = False

                    if match_filter:
                        if current_section == "ENGINEER":
                            rows_engineer.append((row_idx, row_vals, row_str))
                        else:
                            rows_pengelola.append((row_idx, row_vals, row_str))

                def render_section_title(title, rows_list):
                    if not rows_list:
                        return ""
                    actual_data_count = sum(1 for r_i, r_v, r_s in rows_list if not (("NO" in [x.upper() for x in r_v]) and ("CSE" in [x.upper() for x in r_v])))
                    if is_filtering and actual_data_count == 0:
                        return ""

                    sec_html = f'<div class="excel-container"><table class="excel-table"><tbody>'
                    sec_html += f'<tr><td colspan="{len(valid_col_indices)}" class="row-section-title">{title}</td></tr>'

                    for row_idx, row_vals, row_str in rows_list:
                        is_header_row = ("NO" in [v.upper() for v in row_vals]) and ("CSE" in [v.upper() for v in row_vals] or "PKT" in [v.upper() for v in row_vals])
                        
                        if is_header_row:
                            sec_html += '<tr class="header-row">'
                            for col_idx in valid_col_indices:
                                v = row_vals[col_idx] if col_idx < len(row_vals) else ""
                                if v in ["#DIV/0!", "nan", "None"]: v = ""
                                try:
                                    num_v = float(v)
                                    if num_v.is_integer(): v = str(int(num_v))
                                except Exception: pass
                                cls = "col-no" if col_idx == 0 else ("col-cse" if col_idx == 1 else "")
                                sec_html += f'<td class="{cls}">{v}</td>'
                            sec_html += "</tr>"
                            continue

                        is_total = "TOTAL" in row_str
                        is_id_tidak_tercapai = "ID TIDAK TERCAPAI" in row_str
                        tr_class = ' class="row-summary"' if (is_total or is_id_tidak_tercapai) else ""
                        sec_html += f"<tr{tr_class}>"

                        for col_idx in valid_col_indices:
                            val = row_vals[col_idx] if col_idx < len(row_vals) else ""
                            if val == "" or val in ["#DIV/0!", "nan", "None"]:
                                sec_html += "<td></td>"
                                continue

                            cell_cls = []
                            if col_idx == 0: cell_cls.append("col-no")
                            elif col_idx == 1: cell_cls.append("col-cse")

                            display_val = val
                            try:
                                num_val = float(val)
                                if is_id_tidak_tercapai:
                                    display_val = str(int(round(num_val)))
                                    cell_cls.append("bg-red")
                                elif 0 < num_val <= 100 and col_idx >= 3:
                                    display_val = f"{num_val:.2f}"
                                    if num_val >= 99.20: cell_cls.append("bg-green")
                                    else: cell_cls.append("bg-red")
                                elif num_val.is_integer(): display_val = str(int(num_val))
                                else: display_val = f"{num_val:.2f}"
                            except Exception: pass

                            cls_str = f' class="{" ".join(cell_cls)}"' if cell_cls else ""
                            sec_html += f"<td{cls_str}>{display_val}</td>"
                        sec_html += "</tr>"
                    sec_html += "</tbody></table></div>"
                    return sec_html

                html += render_section_title("ENGINEER", rows_engineer)
                if not is_filtering and rows_pengelola:
                    html += render_section_title("PENGELOLA", rows_pengelola)
                html += "</div>"
                return html

            table_html = build_uptime_daily_html(df_sheet2, allowed_months, selected_daily_cse)
            st.components.v1.html(table_html, height=1400, scrolling=True)
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data Uptime Harian: {e}")

# =============================================================
# HALAMAN 3: UPTIME CSE BY TIPE MESIN
# =============================================================
elif menu_option == "Uptime CSE by Tipe Mesin":
    col_t, col_b = st.columns([5, 1])
    with col_t:
        st.markdown(
            f"<h4 style='margin-bottom: 12px; color: #2c3e50; font-weight: 600;'>📊 Uptime CSE by Tipe Mesin ({selected_month}) - CSE: <span style='color: #3498db;'>{selected_cse}</span></h4>",
            unsafe_allow_html=True,
        )
    with col_b:
        st.write("")
        st.button("⬅️ Back", key="btn_back_cse", use_container_width=True, on_click=go_back_history)

    try:
        table_name_sql = f"uptime_cse_by_tipe_mesin_{selected_month.replace('-', '_')}"
        current_dir = os.path.dirname(os.path.abspath(__file__))
        db_path = None
        try:
            all_files = os.listdir(current_dir)
            master_dbs = [f for f in all_files if 'MASTER' in f and '_db' in f]
            if master_dbs:
                master_dbs.sort()
                db_path = os.path.join(current_dir, master_dbs[-1])
            elif 'atm_database' in all_files:
                db_path = os.path.join(current_dir, 'atm_database')
        except Exception:
            pass

        df_a15 = pd.DataFrame()
        if db_path and os.path.exists(db_path):
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name_sql}'")
                table_exists = cursor.fetchone()
                if table_exists:
                    df_a15 = pd.read_sql(f"SELECT * FROM {table_name_sql}", conn)
                conn.close()
            except Exception:
                pass

        if df_a15.empty:
            xls = pd.ExcelFile(target_excel_file)
            target_sheet = None
            for s in xls.sheet_names:
                if s.upper().strip() in ["BALI NUSRA", "UPTIME HARIAN", "UPTIME CSE"]:
                    target_sheet = s
                    break

            if target_sheet:
                df_a15 = pd.read_excel(xls, target_sheet, skiprows=14, header=None)
                if db_path and os.path.exists(db_path):
                    try:
                        conn = sqlite3.connect(db_path)
                        df_a15.to_sql(table_name_sql, conn, if_exists='replace', index=False)
                        conn.close()
                    except Exception:
                        pass

        if not df_a15.empty:
            def build_filtered_cse_excel_html(df, filter_cse):
                html = """
                <style>
                    .excel-container { max-height: 750px; overflow-y: auto; overflow-x: auto; border: 1px solid #7f8c8d; border-radius: 2px; background-color: #ffffff; }
                    .excel-table { width: 100%; border-collapse: collapse; font-size: 11px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
                    .excel-table td, .excel-table th { border: 1px solid #a6a6a6; padding: 4px 6px; white-space: nowrap; }
                    .col-no { width: 32px !important; text-align: center; }
                    .col-cse { width: 160px !important; text-align: left; padding-left: 6px !important; }
                    .header-excel { background-color: #8ea9db !important; color: #000000 !important; font-weight: bold; text-align: center; position: sticky; top: 0; z-index: 2; }
                    .title-green-excel { background-color: #c6efce !important; color: #000000 !important; font-weight: bold; text-align: center; font-size: 12px; }
                    .row-total { background-color: #d9e1f2 !important; font-weight: bold; }
                    .text-cse-red { color: #c00000 !important; font-weight: bold !important; }
                    .text-cse-black { color: #000000 !important; font-weight: normal !important; }
                    .text-red { color: #c00000 !important; font-weight: bold; }
                    .text-green { color: #008000 !important; font-weight: bold; }
                    .bg-pink-red { background-color: #ffc7ce !important; color: #9c0006 !important; font-weight: bold; text-align: center; }
                    .bg-light-green { background-color: #c6efce !important; color: #006100 !important; font-weight: bold; text-align: center; }
                    .center { text-align: center; }
                </style>
                <div class="excel-container">
                <table class="excel-table"><tbody>
                """

                is_filtering = (filter_cse and filter_cse != "-- Semua CSE --")
                current_table_target_ut = 99.20
                current_table_target_dt = 0.80

                for _, row in df.iterrows():
                    row_vals = ["" if pd.isnull(x) else str(x).strip() for x in row.values]
                    if not any(row_vals): continue
                    row_str = " ".join([v.upper() for v in row_vals])

                    if "UPTIME" in row_str:
                        if "ATM" in row_str:
                            current_table_target_ut = 99.75
                            current_table_target_dt = 0.25
                        else:
                            current_table_target_ut = 99.20
                            current_table_target_dt = 0.80

                        if is_filtering and ("PROVINSI" in row_str or "PRIVINSI" in row_str or "SERVICE AREA" in row_str):
                            continue
                        
                        title_text = [v for v in row_vals if v != ""][0]
                        html += f'<tr><td colspan="9" class="title-green-excel">{title_text}</td></tr>'
                        continue

                    if "NO" in row_vals and ("CSE" in row_vals or "PROVINSI" in row_vals or "SERVICE AREA" in row_str):
                        if is_filtering and ("PROVINSI" in row_str or "PRIVINSI" in row_str or "SERVICE AREA" in row_str):
                            continue
                        html += '<tr class="header-excel">'
                        for col_idx, v in enumerate(row_vals[:9]):
                            cls = "col-no" if col_idx == 0 else ("col-cse" if col_idx == 1 else "center")
                            if col_idx >= 3:
                                try:
                                    f_val = float(v)
                                    if f_val.is_integer(): v = str(int(f_val))
                                except Exception: pass
                            html += f'<td class="{cls}">{v}</td>'
                        html += "</tr>"
                        continue

                    is_summary_or_region_row = any(k in row_str for k in ["PROVINSI", "PRIVINSI", "SERVICE AREA"])
                    is_total_row = "TOTAL" in row_str

                    match_cse_row = False
                    if len(row_vals) > 1:
                        cse_val_cell = str(row.iloc[1]).strip()
                        if is_filtering and cse_val_cell.upper() == filter_cse.upper():
                            match_cse_row = True

                    if is_filtering:
                        if is_summary_or_region_row or is_total_row: continue
                        if not match_cse_row: continue

                    achieve_ut_val = None
                    try:
                        if len(row_vals) > 7 and row_vals[7] not in ["", "#DIV/0!", "nan", "None"]:
                            achieve_ut_val = float(row_vals[7])
                    except Exception: pass

                    is_cse_red = (achieve_ut_val is not None) and (achieve_ut_val < current_table_target_ut)
                    tr_class = ' class="row-total"' if is_total_row else ""
                    html += f"<tr{tr_class}>"

                    for col_idx, val in enumerate(row_vals[:9]):
                        if val == "" or val in ["#DIV/0!", "nan", "None"]:
                            html += "<td></td>"
                            continue

                        cell_cls = []
                        if col_idx == 0: cell_cls.append("col-no")
                        elif col_idx == 1:
                            cell_cls.append("col-cse")
                            if not is_total_row:
                                if is_cse_red: cell_cls.append("text-cse-red")
                                else: cell_cls.append("text-cse-black")
                        else: cell_cls.append("center")

                        display_val = val
                        try:
                            num_val = float(val)
                            if col_idx == 7:
                                display_val = f"{num_val:.2f}"
                                cell_cls.append("bg-light-green" if num_val >= current_table_target_ut else "bg-pink-red")
                            elif col_idx == 8:
                                display_val = f"{num_val:.2f}"
                                if current_table_target_dt == 0.25:
                                    cell_cls.append("text-red" if num_val > 0.25 else "text-green")
                                else:
                                    cell_cls.append("text-red" if num_val > 0.80 else "text-green")
                            elif num_val.is_integer(): display_val = str(int(num_val))
                            else: display_val = f"{num_val:.2f}"
                        except Exception: pass

                        cls_str = f' class="{" ".join(cell_cls)}"' if cell_cls else ""
                        html += f"<td{cls_str}>{display_val}</td>"
                    html += "</tr>"
                html += "</tbody></table></div>"
                return html

            table_html = build_filtered_cse_excel_html(df_a15, selected_cse)
            st.components.v1.html(table_html, height=750, scrolling=True)
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data Uptime CSE by Tipe Mesin: {e}")

# =============================================================
# HALAMAN 4: UPTIME PKT BY TIPE MESIN
# =============================================================
elif menu_option == "Uptime PKT by Tipe Mesin":
    col_t, col_b = st.columns([5, 1])
    with col_t:
        st.markdown(
            f"<h4 style='margin-bottom: 12px; color: #2c3e50; font-weight: 600;'>📊 Uptime PKT by Tipe Mesin ({selected_month}) - PKT: <span style='color: #3498db;'>{selected_pkt}</span></h4>",
            unsafe_allow_html=True,
        )
    with col_b:
        st.write("")
        st.button("⬅️ Back", key="btn_back_pkt", use_container_width=True, on_click=go_back_history)

    try:
        table_name_sql = f"uptime_pkt_by_tipe_mesin_{selected_month.replace('-', '_')}"
        current_dir = os.path.dirname(os.path.abspath(__file__))
        db_path = None
        try:
            all_files = os.listdir(current_dir)
            master_dbs = [f for f in all_files if 'MASTER' in f and '_db' in f]
            if master_dbs:
                master_dbs.sort()
                db_path = os.path.join(current_dir, master_dbs[-1])
            elif 'atm_database' in all_files:
                db_path = os.path.join(current_dir, 'atm_database')
        except Exception:
            pass

        df_pkt_raw = pd.DataFrame()
        if db_path and os.path.exists(db_path):
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name_sql}'")
                table_exists = cursor.fetchone()
                if table_exists:
                    df_pkt_raw = pd.read_sql(f"SELECT * FROM {table_name_sql}", conn)
                conn.close()
            except Exception:
                pass

        if df_pkt_raw.empty:
            xls = pd.ExcelFile(target_excel_file)
            sheet_4_name = "PKT TIPE MESIN"
            if sheet_4_name not in xls.sheet_names:
                if len(xls.sheet_names) >= 4:
                    sheet_4_name = xls.sheet_names[3]

            df_pkt_raw = pd.read_excel(xls, sheet_name=sheet_4_name, header=None)
            if db_path and os.path.exists(db_path):
                try:
                    conn = sqlite3.connect(db_path)
                    df_pkt_raw.to_sql(table_name_sql, conn, if_exists='replace', index=False)
                    conn.close()
                except Exception:
                    pass

        if not df_pkt_raw.empty:
            def build_filtered_pkt_excel_html(df, filter_pkt):
                html = """
                <style>
                    .excel-container { max-height: 750px; overflow-y: auto; overflow-x: auto; border: 1px solid #7f8c8d; border-radius: 2px; background-color: #ffffff; }
                    .excel-table { width: 100%; border-collapse: collapse; font-size: 11px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
                    .excel-table td, .excel-table th { border: 1px solid #a6a6a6; padding: 4px 6px; white-space: nowrap; }
                    .col-no { width: 32px !important; text-align: center; }
                    .col-pkt { width: 180px !important; text-align: left; padding-left: 6px !important; }
                    .header-excel { background-color: #8ea9db !important; color: #000000 !important; font-weight: bold; text-align: center; position: sticky; top: 0; z-index: 2; }
                    .title-green-excel { background-color: #c6efce !important; color: #000000 !important; font-weight: bold; text-align: center; font-size: 12px; }
                    .row-total { background-color: #d9e1f2 !important; font-weight: bold; }
                    .text-pkt-red { color: #c00000 !important; font-weight: bold !important; }
                    .text-pkt-black { color: #000000 !important; font-weight: normal !important; }
                    .text-red { color: #c00000 !important; font-weight: bold; }
                    .text-green { color: #008000 !important; font-weight: bold; }
                    .bg-pink-red { background-color: #ffc7ce !important; color: #9c0006 !important; font-weight: bold; text-align: center; }
                    .bg-light-green { background-color: #c6efce !important; color: #006100 !important; font-weight: bold; text-align: center; }
                    .center { text-align: center; }
                </style>
                <div class="excel-container">
                <table class="excel-table"><tbody>
                """

                is_filtering = (filter_pkt and filter_pkt != "-- Semua PKT --")
                current_table_target_dt = 0.80

                for _, row in df.iterrows():
                    row_vals = ["" if pd.isnull(x) else str(x).strip() for x in row.values]
                    if not any(row_vals): continue
                    row_str = " ".join([v.upper() for v in row_vals])

                    if "UPTIME" in row_str:
                        if "ATM" in row_str: current_table_target_dt = 0.25
                        else: current_table_target_dt = 0.80

                        if is_filtering and ("PROVINSI" in row_str or "PRIVINSI" in row_str or "SERVICE AREA" in row_str):
                            continue
                        title_text = [v for v in row_vals if v != ""][0]
                        html += f'<tr><td colspan="9" class="title-green-excel">{title_text}</td></tr>'
                        continue

                    if "NO" in row_vals and ("PKT" in row_vals or "CSE" in row_vals or "PROVINSI" in row_str):
                        if is_filtering and ("PROVINSI" in row_str or "PRIVINSI" in row_str or "SERVICE AREA" in row_str):
                            continue
                        html += '<tr class="header-excel">'
                        for col_idx, v in enumerate(row_vals[:9]):
                            cls = "col-no" if col_idx == 0 else ("col-pkt" if col_idx == 1 else "center")
                            if col_idx >= 3:
                                try:
                                    f_val = float(v)
                                    if f_val.is_integer(): v = str(int(f_val))
                                except Exception: pass
                            html += f'<td class="{cls}">{v}</td>'
                        html += "</tr>"
                        continue

                    is_summary_or_region_row = any(k in row_str for k in ["PROVINSI", "PRIVINSI", "SERVICE AREA"])
                    is_total_row = "TOTAL" in row_str

                    match_pkt_row = False
                    if len(row_vals) > 1:
                        pkt_val_cell = str(row.iloc[1]).strip()
                        if is_filtering and pkt_val_cell.upper() == filter_pkt.upper():
                            match_pkt_row = True

                    if is_filtering:
                        if is_summary_or_region_row or is_total_row: continue
                        if not match_pkt_row: continue

                    achieve_ut_val = None
                    try:
                        if len(row_vals) > 7 and row_vals[7] not in ["", "#DIV/0!", "nan", "None"]:
                            achieve_ut_val = float(row_vals[7])
                    except Exception: pass

                    is_pkt_red = (achieve_ut_val is not None) and (achieve_ut_val < 99.20)
                    tr_class = ' class="row-total"' if is_total_row else ""
                    html += f"<tr{tr_class}>"

                    for col_idx, val in enumerate(row_vals[:9]):
                        if val == "" or val in ["#DIV/0!", "nan", "None"]:
                            html += "<td></td>"
                            continue

                        cell_cls = []
                        if col_idx == 0: cell_cls.append("col-no")
                        elif col_idx == 1:
                            cell_cls.append("col-pkt")
                            if not is_total_row:
                                if is_pkt_red: cell_cls.append("text-pkt-red")
                                else: cell_cls.append("text-pkt-black")
                        else: cell_cls.append("center")

                        display_val = val
                        try:
                            num_val = float(val)
                            if col_idx == 7:
                                display_val = f"{num_val:.2f}"
                                cell_cls.append("bg-light-green" if num_val >= 99.20 else "bg-pink-red")
                            elif col_idx == 8:
                                display_val = f"{num_val:.2f}"
                                if current_table_target_dt == 0.25:
                                    cell_cls.append("text-red" if num_val > 0.25 else "text-green")
                                else:
                                    cell_cls.append("text-red" if num_val > 0.80 else "text-green")
                            elif num_val.is_integer(): display_val = str(int(num_val))
                            else: display_val = f"{num_val:.2f}"
                        except Exception: pass

                        cls_str = f' class="{" ".join(cell_cls)}"' if cell_cls else ""
                        html += f"<td{cls_str}>{display_val}</td>"
                    html += "</tr>"
                html += "</tbody></table></div>"
                return html

            table_html = build_filtered_pkt_excel_html(df_pkt_raw, selected_pkt)
            st.components.v1.html(table_html, height=750, scrolling=True)
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses sheet PKT TIPE MESIN: {e}")

# =============================================================
# HALAMAN 5: RIWAYAT KUNJUNGAN
# =============================================================
elif menu_option == "Riwayat Kunjungan (Visit)":
    col_title, col_back = st.columns([5, 1])

    with col_title:
        st.markdown(
            f"<h4 style='margin-bottom: 12px; color: #2c3e50; font-weight: 600;'>📋 Riwayat Kunjungan WSID: <span style='color: #3498db;'>{selected_wsid}</span> ({selected_month})</h4>",
            unsafe_allow_html=True,
        )

    with col_back:
        st.write("")
        st.button(
            "⬅️ Back",
            key="btn_back_visit",
            use_container_width=True,
            on_click=go_back_history,
        )

    try:
        table_name_sql = f"riwayat_kunjungan_{selected_month.replace('-', '_')}"
        current_dir = os.path.dirname(os.path.abspath(__file__))
        db_path = None
        try:
            all_files = os.listdir(current_dir)
            master_dbs = [f for f in all_files if 'MASTER' in f and '_db' in f]
            if master_dbs:
                master_dbs.sort()
                db_path = os.path.join(current_dir, master_dbs[-1])
            elif 'atm_database' in all_files:
                db_path = os.path.join(current_dir, 'atm_database')
        except Exception:
            pass

        df_visit = pd.DataFrame()
        if db_path and os.path.exists(db_path):
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name_sql}'")
                table_exists = cursor.fetchone()
                if table_exists:
                    df_visit = pd.read_sql(f"SELECT * FROM {table_name_sql}", conn)
                conn.close()
            except Exception:
                pass

        if df_visit.empty:
            xls = pd.ExcelFile(target_excel_file)
            sheet_visit_name = None
            for s in xls.sheet_names:
                if "VISIT" in s.upper() or "KUNJUNGAN" in s.upper() or "HISTORI" in s.upper():
                    sheet_visit_name = s
                    break
            if not sheet_visit_name and len(xls.sheet_names) >= 5:
                sheet_visit_name = xls.sheet_names[4]

            if sheet_visit_name:
                df_visit = pd.read_excel(xls, sheet_name=sheet_visit_name)
                if db_path and os.path.exists(db_path):
                    try:
                        conn = sqlite3.connect(db_path)
                        df_visit.to_sql(table_name_sql, conn, if_exists='replace', index=False)
                        conn.close()
                    except Exception:
                        pass

        if df_visit.empty:
            st.warning(f"Tidak ada histori kunjungan yang tercatat untuk bulan {selected_month}.")
        else:
            col_id_name = "ID" if "ID" in df_visit.columns else ("WSID" if "WSID" in df_visit.columns else None)
            if col_id_name:
                df_filtered_visit = df_visit[df_visit[col_id_name].astype(str).str.upper() == str(selected_wsid).upper()]

                if df_filtered_visit.empty:
                    st.info(f"Tidak ada histori kunjungan untuk WSID **{selected_wsid}** pada bulan {selected_month}.")
                else:
                    if "STARTED" in df_filtered_visit.columns:
                        df_filtered_visit = df_filtered_visit.sort_values(by="STARTED", ascending=True)

                    visit_rows_html = []
                    for _, r in df_filtered_visit.iterrows():
                        ticket = str(r.get("TICKET", "-")) if pd.notnull(r.get("TICKET")) else "-"
                        svc_type = str(r.get("SVC_TYPE", "-")) if pd.notnull(r.get("SVC_TYPE")) else "-"
                        started = str(r.get("STARTED", "-")) if pd.notnull(r.get("STARTED")) else "-"
                        finished = str(r.get("FINISHED", "-")) if pd.notnull(r.get("FINISHED")) else "-"
                        spart = str(r.get("SPART", "-")) if pd.notnull(r.get("SPART")) else "-"
                        solution = str(r.get("SOLUTION", "-")) if pd.notnull(r.get("SOLUTION")) else "-"

                        visit_rows_html.append(f"""
                        <tr>
                            <td style="font-weight:600; vertical-align: top; border: 1px solid #dcdfe6; padding: 10px 12px;">{ticket}</td>
                            <td style="font-weight:600; vertical-align: top; border: 1px solid #dcdfe6; padding: 10px 12px;">{svc_type}</td>
                            <td style="vertical-align: top; border: 1px solid #dcdfe6; padding: 10px 12px;">{started}</td>
                            <td style="vertical-align: top; border: 1px solid #dcdfe6; padding: 10px 12px;">{finished}</td>
                            <td style="vertical-align: top; border: 1px solid #dcdfe6; padding: 10px 12px;">{spart}</td>
                            <td style="text-align: left; line-height: 1.4; vertical-align: top; border: 1px solid #dcdfe6; padding: 10px 12px;">{solution}</td>
                        </tr>
                        """)

                    body_html = "".join(visit_rows_html)
                    visit_table_html = f"""
                    <!DOCTYPE html>
                    <html lang="id">
                    <head>
                        <style>
                            * {{ box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; }}
                            body {{ background-color: #f4f6f9; padding: 10px; }}
                            .visit-card {{ background: #ffffff; border-radius: 6px; overflow: hidden; box-shadow: 0 2px 6px rgba(0,0,0,0.1); border: 1px solid #dcdfe6; }}
                            .visit-header {{ background-color: #3598db; color: #ffffff; padding: 10px 15px; font-weight: bold; font-size: 14px; display: flex; justify-content: space-between; align-items: center; }}
                            .visit-table-container {{ max-height: 700px; overflow-y: auto; overflow-x: auto; }}
                            .visit-table {{ width: 100%; border-collapse: collapse; font-size: 12px; color: #333333; }}
                            .visit-table th {{ background-color: #8ea9db; color: #000000; font-weight: bold; text-align: left; padding: 10px 12px; border: 1px solid #dcdfe6; position: sticky; top: 0; z-index: 1; }}
                            .visit-table td {{ border: 1px solid #dcdfe6; padding: 10px 12px; }}
                            .visit-table tr:hover {{ background-color: #f1f5f9; }}
                        </style>
                    </head>
                    <body>
                        <div class="visit-card">
                            <div class="visit-header">
                                <span>Riwayat Kunjungan</span>
                                <span>▼</span>
                            </div>
                            <div class="visit-table-container">
                                <table class="visit-table">
                                    <thead>
                                        <tr>
                                            <th style="width: 13%;">TICKET</th>
                                            <th style="width: 9%;">SVC_TYPE</th>
                                            <th style="width: 15%;">STARTED</th>
                                            <th style="width: 15%;">FINISHED</th>
                                            <th style="width: 12%;">SPART</th>
                                            <th style="width: 36%;">SOLUTION</th>
                                        </tr>
                                    </thead>
                                    <tbody>{body_html}</tbody>
                                </table>
                            </div>
                        </div>
                    </body>
                    </html>
                    """
                    st.components.v1.html(visit_table_html, height=750, scrolling=True)
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses Riwayat Kunjungan: {e}")

# =============================================================
# HALAMAN 6: TERITORI MESIN ENGINEER
# =============================================================
elif menu_option == "Teritori Mesin Engineer":
    db_name = "atm_database.db"
    table_teritori_name = f"teritori_mesin_{selected_year}_{selected_month_num}"
    
    def get_or_create_teritori_data():
        conn = sqlite3.connect(db_name)
        query_check = f"SELECT name FROM sqlite_master WHERE type='table' and name='{table_teritori_name}';"
        df_check = pd.read_sql(query_check, conn)
        
        if not df_check.empty:
            df_res = pd.read_sql(f"SELECT * FROM {table_teritori_name}", conn)
            conn.close()
            return df_res
            
        target_excel_file = f"MASTER_{selected_year}-{selected_month_num}.xlsx"
        if not os.path.exists(target_excel_file):
            matching_masters = glob.glob(f"MASTER_{selected_year}-*.xlsx")
            if matching_masters:
                matching_masters.sort(reverse=True)
                target_excel_file = matching_masters[0]
            elif os.path.exists("MASTER.xlsx"):
                target_excel_file = "MASTER.xlsx"

        df_teritory_local = pd.DataFrame()
        if os.path.exists(target_excel_file):
            try:
                xls_t = pd.ExcelFile(target_excel_file)
                target_sname = xls_t.sheet_names[5] if len(xls_t.sheet_names) >= 6 else None
                for s in xls_t.sheet_names:
                    if "atm total" in s.lower():
                        target_sname = s
                        break
                if target_sname:
                    df_teritory_local = pd.read_excel(xls_t, target_sname, skiprows=5)
            except Exception:
                pass

        all_records = []
        if not df_teritory_local.empty:
            for _, r in df_teritory_local.iterrows():
                group_val = str(r.iloc[10]).strip() if len(r) > 10 and pd.notnull(r.iloc[10]) else ""
                if "BALI NUSRA" in group_val.upper():
                    w_val = str(r.iloc[1]) if len(r) > 1 and pd.notnull(r.iloc[1]) else "-"
                    l_val = str(r.iloc[3]) if len(r) > 3 and pd.notnull(r.iloc[3]) else "-"
                    s_val = str(r.iloc[2]) if len(r) > 2 and pd.notnull(r.iloc[2]) else "-"
                    t_val = str(r.iloc[38]) if len(r) > 38 and pd.notnull(r.iloc[38]) else "-"
                    sa_val = str(r.iloc[9]) if len(r) > 9 and pd.notnull(r.iloc[9]) else "-"
                    sn_val = str(r.iloc[33]) if len(r) > 33 and pd.notnull(r.iloc[33]) else "-"
                    p_val = str(r.iloc[4]) if len(r) > 4 and pd.notnull(r.iloc[4]) else "-"
                    
                    frx_val = r.iloc[13] if len(r) > 13 and pd.notnull(r.iloc[13]) else 0
                    try:
                        frx_float = float(frx_val)
                        frx_formatted = str(int(frx_float)) if frx_float.is_integer() else f"{frx_float:.2f}"
                    except Exception:
                        frx_float = 0.0
                        frx_formatted = "0"

                    u_val = r.iloc[35] if len(r) > 35 and pd.notnull(r.iloc[35]) else None  
                    try:
                        u_float = float(u_val)
                        if u_float <= 1.0: u_float = u_float * 100
                        u_formatted = f"{u_float:.2f}"
                    except Exception:
                        u_float = 0.0
                        u_formatted = "-"

                    if w_val.upper() not in ["WSID", "ID", "NAN", "NONE", "-"] and w_val.strip() != "":
                        all_records.append({
                            "WSID": w_val, "LOKASI": l_val, "SE": s_val, "TYPE": t_val,
                            "SERVICE AREA": sa_val, "SN": sn_val, "PENGELOLA": p_val,
                            "FRX": frx_formatted, "frx_float": frx_float,
                            "UPTIME": u_formatted, "uptime_float": u_float
                        })

        df_rec = pd.DataFrame(all_records)
        if not df_rec.empty:
            df_rec.to_sql(table_teritori_name, conn, if_exists="replace", index=False)
        conn.close()
        return df_rec

    df_rec = get_or_create_teritori_data()

    if st.session_state.get("teritory_detail_wsid"):
        target_clicked_wsid = st.session_state["teritory_detail_wsid"]
        detail_mtype = "CRM"
        detail_freq_dt_val = 0
        detail_achieve_ut_float = 0.0

        match_wsid_rec = df_rec[df_rec["WSID"].astype(str).str.upper() == str(target_clicked_wsid).upper()]
        if not match_wsid_rec.empty:
            row_rec = match_wsid_rec.iloc[0]
            detail_mtype = str(row_rec.get("TYPE", "CRM")).upper()
            detail_freq_dt_val = int(float(row_rec.get("frx_float", 0)))
            detail_achieve_ut_float = float(row_rec.get("uptime_float", 0.0))

        detail_achieve_ut_str = f"{round(detail_achieve_ut_float, 2)}%"
        detail_target_ut = 99.75 if "ATM" in detail_mtype else 99.20
        detail_is_tercapai = detail_achieve_ut_float >= detail_target_ut

        detail_status_text = "Tercapai" if detail_is_tercapai else "Tidak tercapai"
        detail_status_color = "#27ae60" if detail_is_tercapai else "#e74c3c"
        detail_ut_color = detail_status_color
        detail_frx_color = "#e74c3c" if detail_freq_dt_val > 0 else "#27ae60"

        col_banner, col_bck_det = st.columns([4.2, 1.8])
        with col_banner:
            st.markdown(f"""
            <div style="background-color: #ffffff; border-radius: 6px; padding: 10px 15px; border: 1px solid #dcdfe6; margin-bottom: 10px; font-weight: bold; font-size: 14px; color: #2c3e50;">
                UPTIME BULANAN: {detail_mtype}: <span style="color: {detail_ut_color};">{detail_achieve_ut_str}</span> &nbsp;|&nbsp; STATUS: <span style="color: {detail_status_color};">{detail_status_text}</span> &nbsp;|&nbsp; TOTAL FRX: <span style="color: {detail_frx_color};">{detail_freq_dt_val}</span>
            </div>
            """, unsafe_allow_html=True)
        with col_bck_det:
            if st.button("⬅️ Kembali ke Daftar Teritori", key="btn_back_teritory_detail", use_container_width=True):
                st.session_state["teritory_detail_wsid"] = None
                st.rerun()

        col_left_det, col_right_det = st.columns(2)
        with col_left_det:
            st.markdown(f"<h5 style='color: #2c3e50; font-size: 15px;'>Uptime Harian ({selected_month})</h5>", unsafe_allow_html=True)
            year_str, month_str = selected_month.split("-")
            table_frx_rows_html = []
            
            eng_name_for_title = "-"
            loc_name_for_title = "-"

            if not df_site.empty and "ID" in df_site.columns:
                ms = df_site[df_site["ID"].astype(str).str.upper() == str(target_clicked_wsid).upper()]
                if not ms.empty:
                    eng_name_for_title = str(ms.iloc[0].get("SE_NAME", "-"))

            match_loc = df_rec[df_rec["WSID"].astype(str).str.upper() == str(target_clicked_wsid).upper()]
            if not match_loc.empty:
                loc_name_for_title = str(match_loc.iloc[0].get("LOKASI", "-"))

            if daily_sheets:
                sorted_days = sorted(daily_sheets.keys())
                for d in sorted_days:
                    df_d = daily_sheets[d]
                    row_match = df_d[df_d["WSID"].astype(str).str.upper() == str(target_clicked_wsid).upper()]
                    tgl_str = f"{d:02d}/{month_str}/{year_str}"

                    has_valid_day_data = False
                    if not row_match.empty:
                        r = row_match.iloc[0]
                        for col_k in ["FRX", "DUR", "%", "FRX.1", "DUR.1", "%.1"]:
                            val_check = r.get(col_k)
                            if pd.notnull(val_check) and str(val_check).strip() not in ["", "nan", "None"]:
                                has_valid_day_data = True
                                break

                    if has_valid_day_data and not row_match.empty:
                        r = row_match.iloc[0]
                        hw_freq = int(float(r.get("FRX", 0))) if pd.notnull(r.get("FRX")) else 0
                        hw_dur = int(float(r.get("DUR", 0))) if pd.notnull(r.get("DUR")) else 0
                        hw_dt = round(float(r.get("%", 0)), 2) if pd.notnull(r.get("%")) else 0.0

                        p01_freq = int(float(r.get("FRX.1", 0))) if pd.notnull(r.get("FRX.1")) else 0
                        p01_dur = int(float(r.get("DUR.1", 0))) if pd.notnull(r.get("DUR.1")) else 0
                        p01_dt = round(float(r.get("%.1", 0)), 2) if pd.notnull(r.get("%.1")) else 0.0

                        p02_freq = int(float(r.get("FRX.2", 0))) if pd.notnull(r.get("FRX.2")) else 0
                        p02_dur = int(float(r.get("DUR.2", 0))) if pd.notnull(r.get("DUR.2")) else 0
                        p02_dt = round(float(r.get("%.2", 0)), 2) if pd.notnull(r.get("%.2")) else 0.0

                        p03_freq = int(float(r.get("FRX.3", 0))) if pd.notnull(r.get("FRX.3")) else 0
                        p03_dur = int(float(r.get("DUR.3", 0))) if pd.notnull(r.get("DUR.3")) else 0
                        p03_dt = round(float(r.get("%.3", 0)), 2) if pd.notnull(r.get("%.3")) else 0.0

                        p04_freq = int(float(r.get("FRX.4", 0))) if pd.notnull(r.get("FRX.4")) else 0
                        p04_dur = int(float(r.get("DUR.4", 0))) if pd.notnull(r.get("DUR.4")) else 0
                        p04_dt = round(float(r.get("%.4", 0)), 2) if pd.notnull(r.get("%.4")) else 0.0

                        row_bg = "#fcdad7" if (hw_freq > 0 or hw_dur > 0) else "#d5f5e3"
                        dt_color = "#c0392b" if hw_dt > 0 else "#27ae60"

                        row_html = f"""
                        <tr style="background-color: {row_bg}; color: #000000;">
                            <td style="font-weight:bold; text-align:center;">{tgl_str}</td>
                            <td>{hw_freq}</td><td>{hw_dur}</td><td style="color:{dt_color}; font-weight:bold;">{hw_dt}</td>
                            <td>{p01_freq}</td><td>{p01_dur}</td><td>{p01_dt}</td>
                            <td>{p02_freq}</td><td>{p02_dur}</td><td>{p02_dt}</td>
                            <td>{p03_freq}</td><td>{p03_dur}</td><td>{p03_dt}</td>
                            <td>{p04_freq}</td><td>{p04_dur}</td><td>{p04_dt}</td>
                        </tr>
                        """
                    else:
                        row_html = f"""
                        <tr style="background-color: #ffffff; color: #000000;">
                            <td style="font-weight:bold; text-align:center;">{tgl_str}</td>
                            <td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
                        </tr>
                        """
                    table_frx_rows_html.append(row_html)

            table_frx_body = "".join(table_frx_rows_html)
            uptime_harian_box_html = f"""
            <!DOCTYPE html>
            <html lang="id">
            <head>
                <style>
                    * {{ box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; }}
                    .card-box {{ background: #ffffff; border-radius: 6px; overflow: hidden; border: 1px solid #dcdfe6; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
                    .card-head {{ background-color: #7030a0; color: #ffffff; padding: 8px 12px; font-weight: bold; font-size: 13px; display: flex; justify-content: space-between; align-items: center; }}
                    .table-full {{ width: 100%; overflow-x: auto; }}
                    .frx-excel-table {{ width: 100%; border-collapse: collapse; font-size: 10px; text-align: center; }}
                    .frx-excel-table th, .frx-excel-table td {{ border: 1px solid #e2b667; padding: 4px 2px; }}
                    .p02-header-bg {{ background-color: #f39c12 !important; color: #000000 !important; font-size: 11px; font-weight: bold; }}
                    .p02-sub-header {{ background-color: #fdf3e7 !important; color: #000000 !important; font-weight: bold; }}
                </style>
            </head>
            <body>
                <div class="card-box">
                    <div class="card-head">
                        <span>Uptime Harian | {target_clicked_wsid} | {loc_name_for_title} | {eng_name_for_title}</span>
                        <span>▼</span>
                    </div>
                    <div class="table-full">
                        <table class="frx-excel-table">
                            <thead>
                                <tr class="p02-header-bg">
                                    <th rowspan="2">Tgl</th><th colspan="3">HW TOTAL</th><th colspan="3">P01 - Pick Modul</th><th colspan="3">P02 - Presenter</th><th colspan="3">P03 - Reject</th><th colspan="3">P04 - Card Reader</th>
                                </tr>
                                <tr class="p02-sub-header">
                                    <th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th>
                                </tr>
                            </thead>
                            <tbody>{table_frx_body}</tbody>
                        </table>
                    </div>
                </div>
            </body>
            </html>
            """
            st.components.v1.html(uptime_harian_box_html, height=1250, scrolling=False)

        with col_right_det:
            st.markdown(f"<h5 style='color: #2c3e50; font-size: 15px;'>Riwayat Kunjungan ({selected_month})</h5>", unsafe_allow_html=True)
            visit_rows_html = []
            if not df_visit.empty:
                col_id_name = "ID" if "ID" in df_visit.columns else ("WSID" if "WSID" in df_visit.columns else None)
                if col_id_name:
                    df_filtered_visit = df_visit[df_visit[col_id_name].astype(str).str.upper() == str(target_clicked_wsid).upper()]
                    if not df_filtered_visit.empty:
                        if "STARTED" in df_filtered_visit.columns:
                            df_filtered_visit = df_filtered_visit.sort_values(by="STARTED", ascending=True)

                        for _, r in df_filtered_visit.iterrows():
                            ticket = str(r.get("TICKET", "-")) if pd.notnull(r.get("TICKET")) else "-"
                            svc_type = str(r.get("SVC_TYPE", "-")) if pd.notnull(r.get("SVC_TYPE")) else "-"
                            started = str(r.get("STARTED", "-")) if pd.notnull(r.get("STARTED")) else "-"
                            finished = str(r.get("FINISHED", "-")) if pd.notnull(r.get("FINISHED")) else "-"
                            spart = str(r.get("SPART", "-")) if pd.notnull(r.get("SPART")) else "-"
                            solution = str(r.get("SOLUTION", "-")) if pd.notnull(r.get("SOLUTION")) else "-"

                            visit_rows_html.append(f"""
                            <tr>
                                <td style="font-weight:600; vertical-align: top; border: 1px solid #dcdfe6; padding: 8px;">{ticket}</td>
                                <td style="font-weight:600; vertical-align: top; border: 1px solid #dcdfe6; padding: 8px;">{svc_type}</td>
                                <td style="vertical-align: top; border: 1px solid #dcdfe6; padding: 8px; font-size: 10px;">{started}</td>
                                <td style="vertical-align: top; border: 1px solid #dcdfe6; padding: 8px; font-size: 10px;">{finished}</td>
                                <td style="vertical-align: top; border: 1px solid #dcdfe6; padding: 8px;">{spart}</td>
                                <td style="text-align: left; vertical-align: top; border: 1px solid #dcdfe6; padding: 8px;">{solution}</td>
                            </tr>
                            """)

            body_visit_html = "".join(visit_rows_html) if visit_rows_html else '<tr><td colspan="6" style="text-align:center; padding:20px; color:#888;">Tidak ada riwayat kunjungan tercatat pada bulan ini</td></tr>'
            visit_box_html = f"""
            <!DOCTYPE html>
            <html lang="id">
            <head>
                <style>
                    * {{ box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; }}
                    .visit-card {{ background: #ffffff; border-radius: 6px; overflow: hidden; border: 1px solid #dcdfe6; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
                    .visit-header {{ background-color: #3598db; color: #ffffff; padding: 8px 12px; font-weight: bold; font-size: 13px; display: flex; justify-content: space-between; align-items: center; }}
                    .visit-table-container {{ width: 100%; overflow-x: auto; }}
                    .visit-table {{ width: 100%; border-collapse: collapse; font-size: 11px; color: #333333; }}
                    .visit-table th {{ background-color: #8ea9db; color: #000000; font-weight: bold; text-align: left; padding: 8px; border: 1px solid #dcdfe6; position: sticky; top: 0; z-index: 1; }}
                    .visit-table td {{ border: 1px solid #dcdfe6; padding: 8px; }}
                </style>
            </head>
            <body>
                <div class="visit-card">
                    <div class="visit-header">
                        <span>Riwayat Kunjungan</span>
                        <span>▼</span>
                    </div>
                    <div class="visit-table-container">
                        <table class="visit-table">
                            <thead>
                                <tr>
                                    <th style="width: 14%;">TICKET</th>
                                    <th style="width: 10%;">SVC_TYPE</th>
                                    <th style="width: 15%;">STARTED</th>
                                    <th style="width: 15%;">FINISHED</th>
                                    <th style="width: 11%;">SPART</th>
                                    <th style="width: 35%;">SOLUTION</th>
                                </tr>
                            </thead>
                            <tbody>{body_visit_html}</tbody>
                        </table>
                    </div>
                </div>
            </body>
            </html>
            """
            st.components.v1.html(visit_box_html, height=1250, scrolling=False)

    else:
        if not df_rec.empty:
            st.markdown("""
                <style>
                    div[data-testid="stForm"] button[kind="secondary"],
                    div[data-testid="stForm"] button[kind="primary"],
                    div[data-testid="stForm"] button {
                        background-color: #7030a0 !important;
                        color: white !important;
                        border: 1px solid #7030a0 !important;
                    }
                    div[data-testid="stForm"] button:hover {
                        background-color: #5c2482 !important;
                        border: 1px solid #5c2482 !important;
                        color: white !important;
                    }
                </style>
            """, unsafe_allow_html=True)

            display_table_temp = df_rec[["WSID"]].copy()

            unique_se_opts = ["-- Semua SE --"] + sorted(df_rec["SE"].dropna().unique().tolist())
            unique_type_opts = ["-- Semua Type --"] + sorted(df_rec["TYPE"].dropna().unique().tolist())
            unique_sa_opts = ["-- Semua Service Area --"] + sorted(df_rec["SERVICE AREA"].dropna().unique().tolist())
            unique_pengelola_opts = ["-- Semua Pengelola --"] + sorted(df_rec["PENGELOLA"].dropna().unique().tolist())
            unique_sn_opts = ["-- Semua SN --"] + sorted(df_rec["SN"].dropna().unique().tolist())

            with st.form(key="form_teritory_combined"):
                col_wsid, col_btn_go, col_space = st.columns([1.5, 0.6, 3.9])
                with col_wsid:
                    quick_wsid_choice = st.selectbox("Pilih WSID:", options=["-- Pilih WSID --"] + display_table_temp["WSID"].tolist(), label_visibility="collapsed")
                with col_btn_go:
                    submit_button = st.form_submit_button("🚀 Buka", use_container_width=True)

                expander_filter = st.expander("🔍 **Klik Filter Data Teritori**", expanded=False)
                
                with expander_filter:
                    f_col1, f_col_frx, f_col2, f_col3, f_col4, f_col5 = st.columns(6)
                    
                    with f_col1:
                        selected_se = st.selectbox("Filter SE:", options=unique_se_opts, key="form_se")
                    with f_col_frx:
                        selected_frx_sort = st.selectbox("Filter FRX:", options=["-- Pilih Urutan --", "FRX Tertinggi", "FRX Terendah"], key="form_frx")
                    with f_col2:
                        selected_type = st.selectbox("Filter TYPE:", options=unique_type_opts, key="form_type")
                    with f_col3:
                        selected_sa = st.selectbox("Filter SERVICE AREA:", options=unique_sa_opts, key="form_sa")
                    with f_col4:
                        selected_pengelola = st.selectbox("Filter PENGELOLA:", options=unique_pengelola_opts, key="form_pengelola")
                    with f_col5:
                        selected_sn = st.selectbox("Filter SN:", options=unique_sn_opts, key="form_sn")

                    st.markdown("<br>", unsafe_allow_html=True)
                    col_apply_space, col_apply_btn, col_apply_space2 = st.columns([3, 2, 3])
                    with col_apply_btn:
                        apply_filter_btn = st.form_submit_button("🔄 Terapkan Filter", use_container_width=True)

            if submit_button:
                if quick_wsid_choice != "-- Pilih WSID --":
                    st.session_state["teritory_detail_wsid"] = quick_wsid_choice
                    st.rerun()

            filtered_df = df_rec.copy()
            if selected_se and selected_se != "-- Semua SE --":
                filtered_df = filtered_df[filtered_df["SE"] == selected_se]
            if selected_type and selected_type != "-- Semua Type --":
                filtered_df = filtered_df[filtered_df["TYPE"] == selected_type]
            if selected_sa and selected_sa != "-- Semua Service Area --":
                filtered_df = filtered_df[filtered_df["SERVICE AREA"] == selected_sa]
            if selected_sn and selected_sn != "-- Semua SN --":
                filtered_df = filtered_df[filtered_df["SN"] == selected_sn]
            if selected_pengelola and selected_pengelola != "-- Semua Pengelola --":
                filtered_df = filtered_df[filtered_df["PENGELOLA"] == selected_pengelola]
            
            if selected_frx_sort == "FRX Tertinggi":
                filtered_df = filtered_df.sort_values(by="frx_float", ascending=False)
            elif selected_frx_sort == "FRX Terendah":
                filtered_df = filtered_df.sort_values(by="frx_float", ascending=True)
            else:
                filtered_df = filtered_df.sort_values(by="uptime_float", ascending=True)

            st.markdown("<p style='font-size: 12px; color: #555;'>* Klik baris <strong>WSID</strong> pada tabel atau gunakan pilihan cepat di atas untuk melihat detail Uptime Harian dan Riwayat Kunjungan.</p>", unsafe_allow_html=True)

            display_table = filtered_df[["WSID", "LOKASI", "SE", "TYPE", "SERVICE AREA", "SN", "PENGELOLA", "FRX", "UPTIME"]].copy()
            
            def highlight_teritory(row):
                m_type = str(row.get("TYPE", "")).upper()
                frx_val = row.get("frx_float", 0.0) if "frx_float" in filtered_df.columns else 0.0
                uptime_val = row.get("uptime_float", 0.0) if "uptime_float" in filtered_df.columns else 0.0

                frx_color = "color: #9c0006; font-weight: bold;" if frx_val > 0 else ""
                
                uptime_style = ""
                if "ATM" in m_type:
                    if uptime_val > 99.75: uptime_style = "color: #006100; background-color: #c6efce; font-weight: bold;"
                    else: uptime_style = "color: #9c0006; background-color: #ffc7ce; font-weight: bold;"
                elif "CRM" in m_type:
                    if uptime_val > 99.20: uptime_style = "color: #006100; background-color: #c6efce; font-weight: bold;"
                    else: uptime_style = "color: #9c0006; background-color: #ffc7ce; font-weight: bold;"

                styles = [""] * len(row)
                col_list = list(display_table.columns)
                if "FRX" in col_list: styles[col_list.index("FRX")] = frx_color
                if "UPTIME" in col_list: styles[col_list.index("UPTIME")] = uptime_style
                return styles

            styled_df = display_table.copy()
            styled_df["frx_float"] = filtered_df["frx_float"].values
            styled_df["uptime_float"] = filtered_df["uptime_float"].values

            styled_render = styled_df.style.apply(highlight_teritory, axis=1)

            event = st.dataframe(
                styled_render,
                use_container_width=True,
                hide_index=True,
                height=750,
                selection_mode="single-row",
                on_select="rerun",
                column_config={
                    "frx_float": None,
                    "uptime_float": None,
                    "WSID": st.column_config.TextColumn("WSID"),
                    "LOKASI": st.column_config.TextColumn("LOKASI"),
                    "SE": st.column_config.TextColumn("SE"),
                    "TYPE": st.column_config.TextColumn("TYPE"),
                    "SERVICE AREA": st.column_config.TextColumn("SERVICE AREA"),
                    "SN": st.column_config.TextColumn("SN"),
                    "PENGELOLA": st.column_config.TextColumn("PENGELOLA"),
                    "FRX": st.column_config.TextColumn("FRX"),
                    "UPTIME": st.column_config.TextColumn("UPTIME"),
                }
            )

            if event and "selection" in event and "rows" in event["selection"]:
                selected_rows = event["selection"]["rows"]
                if selected_rows:
                    idx = selected_rows[0]
                    clicked_wsid_val = display_table.iloc[idx]["WSID"]
                    st.session_state["teritory_detail_wsid"] = clicked_wsid_val
                    st.rerun()

# =============================================================
# HALAMAN 7: JADWAL STANDBY CSE
# =============================================================
elif menu_option == "Jadwal Standby CSE":
    c_title, c_year, c_month, c_back = st.columns([2.2, 0.9, 1.1, 0.8])
    
    with c_title:
        st.markdown(
            f"<h4 style='margin-top: -5px; margin-bottom: 0px; color: #2c3e50; font-weight: 600; white-space: nowrap;'>📅 Jadwal Standby CSE ({selected_month})</h4>",
            unsafe_allow_html=True,
        )

    if "standby_selected_year" not in st.session_state:
        st.session_state["standby_selected_year"] = 2026

    with c_year:
        st.session_state["standby_selected_year"] = st.selectbox(
            "Pilih Tahun:",
            options=year_options,
            index=year_options.index(st.session_state["standby_selected_year"]) if st.session_state["standby_selected_year"] in year_options else 0,
            key="standby_year_select"
        )

    current_standby_year = st.session_state["standby_selected_year"]
    standby_file = f"JADWAL STANDBY {current_standby_year}.xlsx"
    if not os.path.exists(standby_file) and os.path.exists("JADWAL STANDBY 2026.xlsx"):
        standby_file = "JADWAL STANDBY 2026.xlsx"

    sheet_names = ["Sheet1"]
    if os.path.exists(standby_file):
        try:
            import openpyxl
            wb_standby = openpyxl.load_workbook(standby_file, data_only=True)
            sheet_names = wb_standby.sheetnames
        except Exception:
            pass

    try:
        active_m_code = selected_month.split("-")[1]
        month_map_full = {
            "01": ["JAN", "JANUARI"], "02": ["FEB", "FEBRUARI"], "03": ["MAR", "MARET"], 
            "04": ["APR", "APRIL"], "05": ["MEI"], "06": ["JUN", "JUNI"], 
            "07": ["JUL", "JULI"], "08": ["AGU", "AGUSTUS"], "09": ["SEP", "SEPTEMBER"], 
            "10": ["OKT", "OKTOBER"], "11": ["NOV", "NOVEMBER"], "12": ["DES", "DESEMBER"]
        }
        target_keywords = month_map_full.get(active_m_code, [])
        default_matched_idx = 0
        for idx, s_name in enumerate(sheet_names):
            s_upper = s_name.upper()
            if any(kw in s_upper for kw in target_keywords):
                default_matched_idx = idx
                break
    except Exception:
        default_matched_idx = 0

    if "standby_sheet_select" not in st.session_state or st.session_state["standby_sheet_select"] not in sheet_names:
        st.session_state["standby_sheet_select"] = sheet_names[default_matched_idx] if default_matched_idx < len(sheet_names) else sheet_names[0]

    with c_month:
        selected_standby_sheet = st.selectbox(
            "Pilih Bulan Standby:", 
            options=sheet_names, 
            index=sheet_names.index(st.session_state["standby_sheet_select"]) if st.session_state["standby_sheet_select"] in sheet_names else default_matched_idx,
            key="standby_sheet_select"
        )

    with c_back:
        st.write("")
        st.button("⬅️ Back", key="btn_back_standby", use_container_width=True, on_click=go_back_history)

    if not os.path.exists(standby_file):
        st.warning(f"File `JADWAL STANDBY {current_standby_year}.xlsx` belum diunggah.")
    else:
        try:
            db_name = "atm_database.db"
            safe_sheet_name = "".join(c if c.isalnum() else "_" for c in selected_standby_sheet)
            table_standby_name = f"standby_{current_standby_year}_{safe_sheet_name}".lower()

            conn = sqlite3.connect(db_name)
            query_check = f"SELECT name FROM sqlite_master WHERE type='table' and name='{table_standby_name}';"
            df_check = pd.read_sql(query_check, conn)

            if not df_check.empty:
                df_saved_html = pd.read_sql(f"SELECT html_content FROM {table_standby_name}", conn)
                conn.close()
                if not df_saved_html.empty:
                    cached_html = df_saved_html.iloc[0]["html_content"]
                    st.components.v1.html(cached_html, height=750, scrolling=True)
            else:
                wb_standby = openpyxl.load_workbook(standby_file, data_only=True)
                ws_st = wb_standby[selected_standby_sheet] if selected_standby_sheet in wb_standby.sheetnames else wb_standby.active

                merged_map = {}
                try:
                    for rng in ws_st.merged_cells.ranges:
                        min_col, min_row, max_col, max_row = rng.min_col, rng.min_row, rng.max_col, rng.max_row
                        merged_map[(min_row, min_col)] = {"rowspan": max_row - min_row + 1, "colspan": max_col - min_col + 1}
                        for r in range(min_row, max_row + 1):
                            for c in range(min_col, max_col + 1):
                                if (r, c) != (min_row, min_col): merged_map[(r, c)] = "SKIP"
                except Exception: pass

                html = """
                <style>
                    .standby-container { max-height: 750px; overflow-y: auto; overflow-x: auto; border: 1px solid #a6a6a6; border-radius: 2px; background-color: #ffffff; }
                    .standby-table { width: 100%; border-collapse: collapse; font-size: 11px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
                    .standby-table td, .standby-table th { border: 1px solid #d9d9d9; padding: 4px 6px; white-space: nowrap; text-align: center; vertical-align: middle; }
                </style>
                <div class="standby-container">
                <table class="standby-table"><tbody>
                """
                
                max_row = ws_st.max_row
                max_col = ws_st.max_column
                
                for r in range(1, max_row + 1):
                    row_vals = [ws_st.cell(row=r, column=c).value for c in range(1, max_col + 1)]
                    if all(v is None for v in row_vals): continue
                
                    html += "<tr>"
                    for c in range(1, max_col + 1):
                        cell_status = merged_map.get((r, c), None)
                        if cell_status == "SKIP": continue
                
                        cell = ws_st.cell(row=r, column=c)
                        val = cell.value if cell.value is not None else ""
                        span_attrs = ""
                        if cell_status and isinstance(cell_status, dict):
                            if cell_status["rowspan"] > 1: span_attrs += f' rowspan="{cell_status["rowspan"]}"'
                            if cell_status["colspan"] > 1: span_attrs += f' colspan="{cell_status["colspan"]}"'

                        bg_style = ""
                        text_color = "#000000"
                        try:
                            if cell.fill and cell.fill.fill_type:
                                fg = cell.fill.fgColor
                                if fg:
                                    f_val = str(getattr(fg, 'value', ''))
                                    if len(f_val) == 8 and f_val.startswith('FF'):
                                        hex_val = "#" + f_val[2:]
                                        bg_style = f"background-color: {hex_val};"
                                        if hex_val.lower() in ["#ff0000", "#00b0f0", "#7030a0", "#002060"]:
                                            text_color = "#ffffff;"
                        except Exception: pass

                        font_bold = "font-weight: bold;" if (cell.font and cell.font.bold) else ""
                        align_style = "text-align: center;"
                        if c in [2, 5]: align_style = "text-align: left; padding-left: 8px;"
                        elif c == 4: align_style = "text-align: left;"
                            
                        custom_style = f"{bg_style} color: {text_color} {font_bold} {align_style}"
                        if r == 1 and c == 1:
                            html += f'<td colspan="{max_col}" style="background-color: #00b0f0; color: #ffffff; font-weight: bold; font-size: 13px; text-align: center; padding: 8px;">{val}</td>'
                            break
                        elif r == 1: continue
                            
                        html += f'<td{span_attrs} style="{custom_style}">{val}</td>'
                    html += "</tr>"
                html += "</tbody></table></div>"

                df_to_save = pd.DataFrame([{"html_content": html}])
                df_to_save.to_sql(table_standby_name, conn, if_exists="replace", index=False)
                conn.close()
                st.components.v1.html(html, height=750, scrolling=True)

        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file Jadwal Standby: {e}")
