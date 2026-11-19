import streamlit as st

import kalkulator_bunker
import noon_report_converter

# Konfigurasi Halaman & Tema Serba Pink
st.set_page_config(page_title="Kalkulator Bunker", page_icon="", layout="wide")

# Custom CSS untuk tema serba pink + Watermark
st.markdown(
    """
    <style>
    /* Background utama */
    .stApp {
        background-color: #fff0f5;
    }
    /* Warna Header */
    h1, h2, h3 {
        color: #d81b60 !important;
    }
    /* Tombol Utama */
    .stButton>button {
        background-color: #ff69b4;
        color: white;
        border-radius: 10px;
        border: none;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #ff1493;
        color: white;
    }
    /* Sidebar / Kotak Login */
    div[data-testid="stForm"] {
        background-color: #ffe4e1;
        padding: 20px;
        border-radius: 15px;
        border: 2px solid #ffb6c1;
    }
    /* Styling */
    .watermark {
        position: fixed;
        bottom: 10px;
        left: 15px;
        font-size: 15px;
        color: #d81b60;
        opacity: 0.6;
        font-weight: bold;
        z-index: 999999;
    }
    </style>

    <!-- Elemen Watermark -->
    <div class="watermark">By: Angelia</div>
""",
    unsafe_allow_html=True,
)

# --- SISTEM LOGIN ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("Login Terlebih Dahulu Ya!")
    with st.form("login_form"):
        password_input = st.text_input("Masukkan Password:", type="password")
        submit_login = st.form_submit_button("Login")

        if submit_login:
            if password_input == "&jel":
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("Password salah!")
    st.stop()

# --- APLIKASI UTAMA (SETELAH LOGIN) ---
# Tombol Logout di sidebar
with st.sidebar:
    st.write("Menu Navigasi")
    if st.button("Logout"):
        st.session_state.logged_in = False
        st.rerun()

tab_kalkulator, tab_converter = st.tabs(["⛽ Kalkulator Bunker", "🚢 Converter Noon Report"])

with tab_kalkulator:
    kalkulator_bunker.render()

with tab_converter:
    noon_report_converter.render()
