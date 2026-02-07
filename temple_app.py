import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import hashlib 
import calendar
import base64
import os
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
        st.error("Supabase credentials not found. Check st.secrets.")
        st.stop()

supabase = get_supabase_client()

# --- CONFIGURATION ---
st.set_page_config(page_title="Temple Management System", layout="wide", page_icon="🕉️")

ADMIN_ROLE = 'admin'
NATCHATHIRAM_OPTIONS = ['Ashwini', 'Bharani', 'Karthigai', 'Rohini', 'Mrigasiram', 'Thiruvathirai', 'Punarpoosam', 'Poosam', 'Ayilyam', 'Magam', 'Poorvam', 'Uthiram', 'Hastham', 'Chithirai', 'Swathi', 'Visakam', 'Anusham', 'Kettai', 'Moolam', 'Pooradam', 'Uthiradam', 'Thiruvonam', 'Avittam', 'Sathayam', 'Poorattathi', 'Uthirattathi', 'Revathi']
MIN_DATE = date(1940, 1, 1)

# TEMPLE DETAILS
TEMPLE_NAME_FULL = "Sree Bhadreshwari Amman Temple Management System"
TRUST_DETAILS = "Samrakshana Seva Trust 174/2004"
ADDRESS_LINE_1 = "Kanjampuram"
ADDRESS_LINE_2 = "Kanniyakumari Dist., Tamil Nadu - 629154"
LOGO_PATH = "amman.jpg" 
BACKGROUND_PATH = "background.jpg"

# --- DB WRAPPER FUNCTIONS ---
def run_supabase_insert(table_name, data):
    try: return supabase.table(table_name).insert(data).execute()
    except Exception as e: st.error(f"Error: {e}"); return None

def get_data(table_name, select="*"):
    try:
        res = supabase.table(table_name).select(select).execute()
        return pd.DataFrame(res.data)
    except: return pd.DataFrame()

# --- UI UTILITIES ---
def get_base64_of_bin_file(bin_file):
    if os.path.exists(bin_file):
        with open(bin_file, 'rb') as f: return base64.b64encode(f.read()).decode()
    return None

def to_excel(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return out.getvalue()

def format_date_for_ui(val):
    try: return pd.to_datetime(val).strftime('%d/%m/%Y')
    except: return str(val)

# --- LOGIN & AUTH ---
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "Home Dashboard"

if not st.session_state.logged_in:
    bg_img = get_base64_of_bin_file(BACKGROUND_PATH)
    st.markdown(f"""
        <style>
        .stApp {{ 
            {"background-image: url('data:image/jpg;base64," + bg_img + "');" if bg_img else ""}
            background-size: 75%; background-position: center; background-repeat: no-repeat;
            background-attachment: fixed; background-color: #bf953f; 
        }}
        .login-box {{ background-color: #800000; color: #FFD700; padding: 20px; text-align: center; border: 2px solid #FFD700; }}
        </style>
    """, unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown("<div style='height: 15vh;'></div>", unsafe_allow_html=True)
        st.markdown("<div class='login-box'><h1>🕉️</h1><h2>Temple Management Login</h2></div>", unsafe_allow_html=True)
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.button("SIGN IN"):
            res = supabase.table('users').select('*').eq('username', u).execute()
            if res.data and res.data[0]['password_hash'] == hashlib.sha256(p.encode()).hexdigest():
                st.session_state.logged_in = True
                st.session_state.username = u
                st.session_state.role = res.data[0]['role']
                st.session_state.rights = res.data[0].get('rights', '').split(',')
                st.rerun()
            else: st.error("Invalid Login")
    st.stop()

# --- MAIN PAGE STYLING (GOLD GRADIENT) ---
st.markdown("""
    <style>
    .stApp { background: linear-gradient(135deg, #bf953f 0%, #fcf6ba 25%, #b38728 50%, #fbf5b7 75%, #aa771c 100%); background-attachment: fixed; }
    div.stButton > button:first-child { background-color: #800000; color: #FFD700; border: 1px solid #FFD700; font-weight: bold; }
    div.stButton > button:hover { background-color: #A00000; color: white; border-color: white; }
    </style>
""", unsafe_allow_html=True)

def render_navigation():
    pages = ["Home Dashboard", "Enroll", "Search", "Billing", "Expenses", "Reports", "Assets", "Samayavakuppu", "Settings"]
    # Filter by user rights
    if st.session_state.role != 'admin':
        pages = [p for p in pages if p in st.session_state.rights]
    
    cols = st.columns(len(pages) + 1)
    for i, p in enumerate(pages):
        if cols[i].button(p):
            st.session_state.current_page = p
            st.rerun()
    if cols[-1].button("LOGOUT"):
        st.session_state.logged_in = False
        st.rerun()
    st.markdown("---")

def page_header():
    c1, c2 = st.columns([1, 8])
    with c1: st.markdown("### 🕉️")
    with c2: st.markdown(f"<h1 style='color: #800000;'>{TEMPLE_NAME_FULL}</h1>", unsafe_allow_html=True)

# --- PAGES ---

if st.session_state.current_page == "Home Dashboard":
    page_header(); render_navigation()
    st.subheader(f"Welcome, {st.session_state.username}")
    
    # HOME SCREEN DETAILS
    df_fam = get_data("families")
    df_trans = get_data("transactions")
    df_exp = get_data("users_expenses")
    
    t_inc = df_trans['amount'].sum() if not df_trans.empty else 0
    t_exp = df_exp['amount'].sum() if not df_exp.empty else 0
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Devotees", len(df_fam))
    m2.metric("Total Income", f"₹ {t_inc:,.2f}")
    m3.metric("Total Expenses", f"₹ {t_exp:,.2f}")
    m4.metric("Net Profit", f"₹ {(t_inc - t_exp):,.2f}")
    
    st.markdown("---")
    st.subheader("Recent Transactions")
    if not df_trans.empty:
        st.dataframe(df_trans.sort_values('date', ascending=False).head(10), use_container_width=True)

elif st.session_state.current_page == "Reports":
    page_header(); render_navigation()
    st.header("Financial & Servicewise Reports")
    
    df_serv = get_data("services")
    if not df_serv.empty:
        with st.container(border=True):
            st.write("### Filter Report")
            c1, c2, c3 = st.columns([2, 1, 1])
            sel_svc = c1.selectbox("Select Service", ["All Services"] + df_serv['service_name'].tolist())
            s_date = c2.date_input("From", date.today() - timedelta(30))
            e_date = c3.date_input("To", date.today())
            
            if st.button("Generate Report", type="primary"):
                df_tr = get_data("transactions")
                if not df_tr.empty:
                    df_tr['dt'] = pd.to_datetime(df_tr['date']).dt.date
                    merged = df_tr.merge(df_serv, left_on='service_id', right_on='id', suffixes=('', '_master'))
                    
                    # Filter
                    mask = (merged['dt'] >= s_date) & (merged['dt'] <= e_date)
                    if sel_svc != "All Services": mask &= (merged['service_name'] == sel_svc)
                    
                    report = merged[mask]
                    if not report.empty:
                        st.metric(f"Collection for {sel_svc}", f"₹ {report['amount'].sum():,.2f}")
                        st.dataframe(report[['dt', 'guest_name', 'service_name', 'amount']], use_container_width=True)
                        st.download_button("Download Excel", to_excel(report), "Report.xlsx")
                    else: st.warning("No data found.")

# --- FOOTER ---
st.markdown("<div style='height: 100px;'></div>", unsafe_allow_html=True)
st.markdown("""
    <style>
    .footer { position: fixed; bottom: 0; left: 0; width: 100%; background: #800000; color: #FFD700; text-align: center; padding: 10px; font-weight: bold; border-top: 2px solid #FFD700; }
    </style>
    <div class='footer'>Developed By : Sai Dharshini Info Solution</div>
""", unsafe_allow_html=True)
