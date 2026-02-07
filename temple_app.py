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
        st.error("Supabase credentials not found. Please check your configuration.")
        st.stop()

supabase = get_supabase_client()

# --- CONFIGURATION & PAGE SETUP ---
st.set_page_config(page_title="Temple Management System", layout="wide", page_icon="🕉️")

ADMIN_ROLE = 'admin'
USER_ROLE = 'user'

# Default Options
NATCHATHIRAM_OPTIONS = ['Ashwini', 'Bharani', 'Karthigai', 'Rohini', 'Mrigasiram', 'Thiruvathirai', 'Punarpoosam', 'Poosam', 'Ayilyam', 'Magam', 'Poorvam', 'Uthiram', 'Hastham', 'Chithirai', 'Swathi', 'Visakam', 'Anusham', 'Kettai', 'Moolam', 'Pooradam', 'Uthiradam', 'Thiruvonam', 'Avittam', 'Sathayam', 'Poorattathi', 'Uthirattathi', 'Revathi']
MIN_DATE = date(1940, 1, 1)

# TEMPLE DETAILS
TEMPLE_NAME_FULL = "Sree Bhadreshwari Amman Temple Management System"
TRUST_DETAILS = "Samrakshana Seva Trust 174/2004"
ADDRESS_LINE_1 = "Kanjampuram"
ADDRESS_LINE_2 = "Kanniyakumari Dist., Tamil Nadu - 629154"
LOGO_PATH = "amman.jpg" 
BACKGROUND_PATH = "background.jpg"

# --- DATABASE WRAPPERS ---
def run_supabase_insert(table_name, data):
    try: return supabase.table(table_name).insert(data).execute()
    except Exception as e: st.error(f"Error: {e}"); return None

def get_data(table_name, select="*"):
    try:
        res = supabase.table(table_name).select(select).execute()
        return pd.DataFrame(res.data)
    except: return pd.DataFrame()

# --- UTILITIES ---
def get_base64_of_bin_file(bin_file):
    if os.path.exists(bin_file):
        with open(bin_file, 'rb') as f: return base64.b64encode(f.read()).decode()
    return None

def to_excel(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return out.getvalue()

# --- AUTHENTICATION ---
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
        .login-card {{ background-color: #800000; color: #FFD700; padding: 30px; border: 3px solid #FFD700; text-align: center; }}
        label {{ color: #800000 !important; font-weight: bold; }}
        </style>
    """, unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown("<div style='height: 15vh;'></div>", unsafe_allow_html=True)
        st.markdown("<div class='login-card'><h1>🕉️</h1><h3>Staff Login</h3></div>", unsafe_allow_html=True)
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.button("SIGN IN"):
            res = supabase.table('users').select('*').eq('username', u).execute()
            if res.data and res.data[0]['password_hash'] == hashlib.sha256(p.encode()).hexdigest():
                st.session_state.logged_in, st.session_state.username = True, u
                st.session_state.role = res.data[0]['role']
                st.session_state.rights = res.data[0].get('rights', '').split(',')
                st.rerun()
            else: st.error("Invalid Credentials")
    st.stop()

# --- MAIN APP UI ---
st.markdown("""
    <style>
    .stApp { background: linear-gradient(135deg, #bf953f 0%, #fcf6ba 25%, #b38728 50%, #fbf5b7 75%, #aa771c 100%); background-attachment: fixed; }
    div.stButton > button { background-color: #800000 !important; color: #FFD700 !important; border: 1px solid #FFD700 !important; font-weight: bold; width: 100%; height: 3.5em; }
    div.stButton > button:hover { background-color: #A00000 !important; color: #FFFFFF !important; }
    </style>
""", unsafe_allow_html=True)

def render_nav():
    all_p = ["Home Dashboard", "Enroll", "Search", "Billing", "Expenses", "Reports", "Assets", "Samayavakuppu", "Settings"]
    display_p = all_p if st.session_state.role == 'admin' else [x for x in all_p if x in st.session_state.rights]
    cols = st.columns(len(display_p) + 1)
    for i, p in enumerate(display_p):
        if cols[i].button(p): st.session_state.current_page = p; st.rerun()
    if cols[-1].button("LOGOUT"): st.session_state.logged_in = False; st.rerun()
    st.markdown("---")

def page_header():
    c1, c2 = st.columns([1, 10])
    with c1: st.markdown("## 🕉️")
    with c2: st.markdown(f"<h1 style='color: #800000; margin-top:-10px;'>{TEMPLE_NAME_FULL}</h1>", unsafe_allow_html=True)

# --- MODULES ---

if st.session_state.current_page == "Home Dashboard":
    page_header(); render_nav()
    st.subheader(f"Namaste, {st.session_state.username.title()}")
    
    # FETCH DATA FOR DASHBOARD
    f_df = get_data("families"); t_df = get_data("transactions"); e_df = get_data("users_expenses")
    inc = t_df['amount'].sum() if not t_df.empty else 0
    exp = e_df['amount'].sum() if not e_df.empty else 0
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Devotees", len(f_df))
    m2.metric("Total Collection", f"₹ {inc:,.2f}")
    m3.metric("Total Expenses", f"₹ {exp:,.2f}")
    m4.metric("Balance", f"₹ {(inc - exp):,.2f}")
    
    st.divider()
    st.subheader("Recent Activity")
    if not t_df.empty: st.dataframe(t_df.sort_values('date', ascending=False).head(5), use_container_width=True, hide_index=True)

elif st.session_state.current_page == "Reports":
    page_header(); render_nav()
    st.header("Financial & Servicewise Reports")
    
    df_serv = get_data("services")
    if not df_serv.empty:
        with st.container(border=True):
            st.markdown("### 🔍 Advanced Filter")
            c1, c2, c3 = st.columns([2, 1, 1])
            
            # 1. Choose Service
            svc_list = ["All Services"] + df_serv['service_name'].tolist()
            target_svc = c1.selectbox("Filter by Service Name", svc_list)
            
            # 2. Choose Date Range
            d_from = c2.date_input("Start Date", date.today() - timedelta(30))
            d_to = c3.date_input("End Date", date.today())
            
            if st.button("Generate Detailed Report", type="primary"):
                df_tr = get_data("transactions")
                if not df_tr.empty:
                    # Clean Dates & Merge Names
                    df_tr['dt_obj'] = pd.to_datetime(df_tr['date']).dt.date
                    merged = df_tr.merge(df_serv, left_on='service_id', right_on='id', suffixes=('', '_master'))
                    
                    # Apply Date Filter
                    mask = (merged['dt_obj'] >= d_from) & (merged['dt_obj'] <= d_to)
                    
                    # Apply Service Name Filter
                    if target_svc != "All Services":
                        mask &= (merged['service_name'] == target_svc)
                    
                    final_rep = merged[mask]
                    
                    if not final_rep.empty:
                        st.divider()
                        st.subheader(f"Report for {target_svc}")
                        
                        r1, r2 = st.columns(2)
                        r1.metric("Collection Amount", f"₹ {final_rep['amount'].sum():,.2f}")
                        r2.metric("Total Bookings", len(final_rep))
                        
                        # Format for display
                        display_df = final_rep[['dt_obj', 'guest_name', 'service_name', 'amount']].copy()
                        display_df.columns = ['Date', 'Devotee Name', 'Service Performed', 'Amount (₹)']
                        st.dataframe(display_df, use_container_width=True, hide_index=True)
                        
                        # Downloads
                        st.download_button("📂 Download Excel Report", to_excel(display_df), f"Report_{target_svc}_{d_from}.xlsx")
                    else:
                        st.warning(f"No transactions found for {target_svc} in this date range.")
                else:
                    st.info("No transaction history found in database.")
    else:
        st.warning("Please configure Services in the Settings menu first.")

# --- OTHER PAGES (Placeholders for your existing logic) ---
elif st.session_state.current_page == "Settings":
    page_header(); render_nav()
    st.subheader("Manage Temple Services")
    with st.form("svc_add"):
        s_name = st.text_input("Service Name (e.g. Pushpanjali)")
        s_price = st.number_input("Standard Rate (₹)", min_value=0.0)
        if st.form_submit_button("Add Service"):
            run_supabase_insert("services", {"service_name": s_name, "price": s_price})
            st.rerun()
    st.dataframe(get_data("services"), use_container_width=True)

# --- FOOTER ---
st.markdown("<div style='height: 100px;'></div>", unsafe_allow_html=True)
st.markdown("""
    <div style='position: fixed; bottom: 0; left: 0; width: 100%; background: #800000; color: #FFD700; text-align: center; padding: 10px; font-weight: bold; border-top: 2px solid #FFD700; z-index: 100;'>
        Developed By : Sai Dharshini Info Solution
    </div>
""", unsafe_allow_html=True)
