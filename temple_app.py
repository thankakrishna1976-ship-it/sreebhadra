import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import hashlib 
import calendar
import base64
import os
import urllib.parse
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader 
from supabase import create_client, Client

# --- SUPABASE CONNECTION SETUP ---
def get_supabase_client() -> Client:
    try:
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
        return create_client(url, key)
    except Exception as e:
        st.error("Supabase credentials not found in st.secrets.")
        st.stop()

supabase = get_supabase_client()

# --- CONFIGURATION ---
st.set_page_config(page_title="Temple Management System", layout="wide", page_icon="🕉️")

ADMIN_ROLE = 'admin'
USER_ROLE = 'user'

# Menu Keys for Rights Management
ALL_MENU_KEYS = [
    "Home Dashboard", "Enroll", "Search", "Billing", 
    "Marriage Bond", "Bond Report", "Expenses", "Reports", "Assets", "Settings"
]

# Standard Options
NATCHATHIRAM_OPTIONS = ['Ashwini', 'Bharani', 'Karthigai', 'Rohini', 'Mrigasiram', 'Thiruvathirai', 'Punarpoosam', 'Poosam', 'Ayilyam', 'Magam', 'Poorvam', 'Uthiram', 'Hastham', 'Chithirai', 'Swathi', 'Visakam', 'Anusham', 'Kettai', 'Moolam', 'Pooradam', 'Uthiradam', 'Thiruvonam', 'Avittam', 'Sathayam', 'Poorattathi', 'Uthirattathi', 'Revathi']
RELATIONSHIP_OPTIONS = ['Wife', 'Son', 'Daughter', 'Mother', 'Father', 'Grand Father', 'Grand Mother', 'Guardian', 'Other']
MIN_DATE = date(1940, 1, 1)

# Temple Info
TEMPLE_NAME_FULL = "Sree Bhadreshwari Amman Temple Management System"
TRUST_DETAILS = "Samrakshana Seva Trust 174/2004"
ADDRESS_LINE_1 = "Kanjampuram"
ADDRESS_LINE_2 = "Kanniyakumari Dist., Tamil Nadu - 629154"
LOGO_PATH = "amman.jpg" 
BACKGROUND_PATH = "background.jpg"

# --- DB WRAPPERS ---
def run_supabase_insert(table_name, data):
    return supabase.table(table_name).insert(data).execute()

def run_supabase_update(table_name, data, row_id):
    return supabase.table(table_name).update(data).eq('id', row_id).execute()

def run_supabase_delete(table_name, row_id):
    return supabase.table(table_name).delete().eq('id', row_id).execute()

def get_data(table_name, select="*"):
    res = supabase.table(table_name).select(select).execute()
    return pd.DataFrame(res.data)

# --- UTILITIES ---
def format_date_for_db(val):
    if pd.isna(val) or not val: return None
    return pd.to_datetime(val).strftime('%Y-%m-%d')

def format_date_for_ui(val):
    if not val or str(val).lower() in ["none", "nat", ""]: return ""
    return pd.to_datetime(val).strftime('%d/%m/%Y')

def safe_date_convert(val):
    if not val or str(val).lower() in ["none", "nat", ""]: return None
    return pd.to_datetime(val).date()

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def get_whatsapp_link(phone, devotee_name, service_name, amount, receipt_id):
    if phone:
        clean_phone = "".join(filter(str.isdigit, str(phone)))
        if len(clean_phone) == 10: clean_phone = f"91{clean_phone}"
        message = f"🙏 *{TEMPLE_NAME_FULL}*\nNamaste *{devotee_name}*,\nReceipt: *#{receipt_id}*\nService: *{service_name}*\nAmount: *₹{amount:,.2f}*\nMay Amman bless you. ✨"
        return f"https://wa.me/{clean_phone}?text={urllib.parse.quote(message)}"
    return None

# --- UI COMPONENTS ---
def render_navigation_bar():
    ALL_PAGES = {
        "Home Dashboard": "🏠 HOME", "Enroll": "📝 ENROLLMENT",
        "Search": "🔍 SEARCH", "Billing": "🧾 BILLING",
        "Marriage Bond": "📜 BOND ENTRY", "Bond Report": "📋 BOND REPORT",
        "Expenses": "💸 EXPENSES", "Reports": "📊 REPORTS",
        "Assets": "🏛️ ASSETS", "Settings": "⚙️ SETTINGS",
    }
    user_rights = st.session_state.get('rights', ["Home Dashboard"])
    if st.session_state.role == ADMIN_ROLE:
        NAV_BAR_PAGES = ALL_PAGES.copy()
        NAV_BAR_PAGES["Users"] = "👥 USERS"
    else:
        NAV_BAR_PAGES = {k: v for k, v in ALL_PAGES.items() if k in user_rights}
    
    cols = st.columns(len(NAV_BAR_PAGES) + 1)
    for i, (key, label) in enumerate(NAV_BAR_PAGES.items()):
        if cols[i].button(label, key=f"nav_{key}", use_container_width=True):
            st.session_state.current_page = key
            st.rerun()
    if cols[-1].button("🚪 LOGOUT", key="nav_logout", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()
    st.write("---")

# --- APP START ---
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "Home Dashboard"

if not st.session_state.logged_in:
    # Simplified login (refer to your previous login logic)
    st.title("Temple Staff Login")
    u, p = st.text_input("User"), st.text_input("Pass", type="password")
    if st.button("Login"):
        res = supabase.table('users').select('*').eq('username', u).execute()
        if res.data and res.data[0]['password_hash'] == hash_password(p):
            st.session_state.logged_in = True
            st.session_state.username = u
            st.session_state.role = res.data[0]['role']
            st.session_state.rights = res.data[0].get('rights', 'Home Dashboard').split(',')
            st.rerun()
    st.stop()

render_navigation_bar()

# --- MODULES ---

if st.session_state.current_page == "Home Dashboard":
    st.title(f"Welcome to {TEMPLE_NAME_FULL}")
    st.info("Select a menu to begin operations.")

elif st.session_state.current_page == "Marriage Bond":
    st.header("📜 Samaya Vakuppu Marriage Bond Entry")
    with st.form("bond_f", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            nm = st.text_input("Student Name *")
            db = st.date_input("DOB", value=None, min_value=MIN_DATE)
            bk = st.text_input("Issuing Bank")
        with c2:
            bn = st.text_input("Bond No *")
            id = st.date_input("Issue Date", value=date.today())
            md = st.date_input("Maturity Date", value=None)
        if st.form_submit_button("Save Bond"):
            if nm and bn:
                run_supabase_insert("marriage_bonds", {"student_name": nm, "dob": format_date_for_db(db), "issuing_bank": bk, "bond_no": bn, "issuing_date": format_date_for_db(id), "maturity_date": format_date_for_db(md)})
                st.success("Record Saved.")
            else: st.error("Name and Bond No required.")

elif st.session_state.current_page == "Bond Report":
    st.header("📋 Marriage Bond Reports")
    df = get_data("marriage_bonds")
    if not df.empty:
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.divider()
        b_sel = st.selectbox("Select Bond to Edit/Delete", df['student_name'].tolist())
        curr = df[df['student_name'] == b_sel].iloc[0]
        with st.form("edit_bond"):
            un = st.text_input("Name", value=curr['student_name'])
            ub = st.text_input("Bank", value=curr['issuing_bank'])
            uno = st.text_input("Bond No", value=curr['bond_no'])
            if st.form_submit_button("Update"):
                run_supabase_update("marriage_bonds", {"student_name": un, "issuing_bank": ub, "bond_no": uno}, curr['id'])
                st.rerun()
        if st.button("🚨 Delete Bond"):
            run_supabase_delete("marriage_bonds", curr['id'])
            st.rerun()

# --- (Other modules: Billing, Enroll, Search etc. remain exactly as your existing logic) ---
# ...
