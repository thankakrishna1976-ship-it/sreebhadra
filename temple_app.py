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
        st.error("Supabase credentials not found in st.secrets. Please check your configuration.")
        st.stop()

supabase = get_supabase_client()

# --- CONFIGURATION & PAGE SETUP ---
st.set_page_config(page_title="Temple Management System", layout="wide", page_icon="🕉️")

ADMIN_ROLE = 'admin'
USER_ROLE = 'user'
GUEST_FAMILY_ID = 0 

# Default Options
DEFAULT_EXPENSE_TYPES = ['Pooja Items', 'Maintenance/Repairs', 'Salary/Dakshina', 'Electricity/Water', 'Annadanam/Food', 'Construction', 'Festivals', 'Administrative', 'Other']
RELATIONSHIP_OPTIONS = ['Wife', 'Son', 'Daughter', 'Mother', 'Father', 'Grand Father', 'Grand Mother', 'Guardian', 'Other']
NATCHATHIRAM_OPTIONS = ['Ashwini', 'Bharani', 'Karthigai', 'Rohini', 'Mrigasiram', 'Thiruvathirai', 'Punarpoosam', 'Poosam', 'Ayilyam', 'Magam', 'Poorvam', 'Uthiram', 'Hastham', 'Chithirai', 'Swathi', 'Visakam', 'Anusham', 'Kettai', 'Moolam', 'Pooradam', 'Uthiradam', 'Thiruvonam', 'Avittam', 'Sathayam', 'Poorattathi', 'Uthirattathi', 'Revathi']

MIN_DATE = date(1940, 1, 1)
MAX_DATE = date(2040, 12, 31)

ALL_MENU_KEYS = ["Home Dashboard", "Enroll", "Search", "Billing", "Expenses", "Reports", "Assets", "Samayavakuppu", "Settings"]

# TEMPLE DETAILS
TEMPLE_NAME_FULL = "Sree Bhadreshwari Amman Temple Management System"
TRUST_DETAILS = "Samrakshana Seva Trust 174/2004"
ADDRESS_LINE_1 = "Kanjampuram"
ADDRESS_LINE_2 = "Kanniyakumari Dist., Tamil Nadu - 629154"
LOGO_PATH = "amman.jpg" 
BACKGROUND_PATH = "background.jpg"

# --- DB WRAPPER FUNCTIONS ---
def run_supabase_insert(table_name, data):
    try:
        response = supabase.table(table_name).insert(data).execute()
        return response
    except Exception as e:
        st.error(f"Database Error: {e}")
        return None

def run_supabase_update(table_name, data, row_id):
    try:
        response = supabase.table(table_name).update(data).eq('id', row_id).execute()
        return response
    except Exception as e:
        st.error(f"Database Error: {e}")
        return None

def run_supabase_delete(table_name, row_id):
    try:
        response = supabase.table(table_name).delete().eq('id', row_id).execute()
        return response
    except Exception as e:
        st.error(f"Database Error: {e}")
        return None

def get_data(table_name, select="*"):
    try:
        response = supabase.table(table_name).select(select).execute()
        return pd.DataFrame(response.data)
    except Exception as e:
        st.error(f"Database Fetch Error: {e}")
        return pd.DataFrame()

# --- UTILITY FUNCTIONS ---
def format_date_for_db(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nat": return None
    try: return pd.to_datetime(val).strftime('%Y-%m-%d')
    except: return None

def format_date_for_ui(val):
    if not val or str(val).lower() in ["none", "nat", ""]: return ""
    try: return pd.to_datetime(val).strftime('%d/%m/%Y')
    except: return str(val)

def safe_date_convert(val):
    if not val or str(val).lower() in ["none", "nat", ""]: return None
    try: return pd.to_datetime(val).date()
    except: return None

def get_base64_of_bin_file(bin_file):
    try:
        if os.path.exists(bin_file):
            with open(bin_file, 'rb') as f: data = f.read()
            return base64.b64encode(data).decode()
    except: pass
    return None

def image_to_base64(image_file):
    if image_file is not None: return base64.b64encode(image_file.getvalue()).decode()
    return ""

def base64_to_image(base64_str):
    if base64_str:
        try: return io.BytesIO(base64.b64decode(base64_str))
        except: return None
    return None

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def verify_user(username, password):
    res = supabase.table('users').select('*').eq('username', username).execute()
    if res.data:
        user_data = res.data[0]
        if user_data['password_hash'] == hash_password(password):
            user_rights = user_data.get('rights', 'Home Dashboard').split(',')
            return True, user_data['role'], user_rights
    return False, None, None

def to_excel(df):
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Report')
        return output.getvalue()
    except: return None

# --- UI RENDERERS ---
def render_footer():
    st.markdown("""
        <style>
        .footer {
            position: fixed; left: 0; bottom: 0; width: 100%;
            background-color: #800000; color: #FFD700; text-align: center;
            padding: 10px 0; font-size: 14px; font-weight: bold;
            letter-spacing: 1px; z-index: 999; border-top: 2px solid #FFD700;
        }
        .main .block-container { padding-bottom: 60px; }
        </style>
        <div class="footer">Developed By : Sai Dharshini Info Solution</div>
    """, unsafe_allow_html=True)

def page_header():
    st.markdown("<style>.stToolbar {visibility: hidden;}</style>", unsafe_allow_html=True)
    col_img, col_title = st.columns([1, 8])
    with col_img:
        try: st.image(LOGO_PATH, width=80)
        except: st.markdown("<h3>🕉️</h3>", unsafe_allow_html=True)
    with col_title:
        st.markdown(f"<h1 style='color: #800000; font-size: 32px; padding-top: 10px; border-bottom: 2px solid #b38728;'>{TEMPLE_NAME_FULL}</h1>", unsafe_allow_html=True)
    st.markdown("---")

def render_navigation_bar():
    ALL_PAGES = {
        "Home Dashboard": {"label": "HOME"}, "Enroll": {"label": "ENROLLMENT"},
        "Search": {"label": "SEARCH"}, "Billing": {"label": "BILLING"},
        "Expenses": {"label": "EXPENSES"}, "Reports": {"label": "REPORTS"},
        "Assets": {"label": "ASSETS"}, "Samayavakuppu": {"label": "SAMAYAVAKUPPU"},
        "Settings": {"label": "SETTINGS"},
    }
    if st.session_state.role == ADMIN_ROLE:
        NAV_BAR_PAGES = ALL_PAGES.copy()
        NAV_BAR_PAGES["Users"] = {"label": "USERS"}
    else:
        user_rights = st.session_state.get('rights', ["Home Dashboard"])
        NAV_BAR_PAGES = {k: v for k, v in ALL_PAGES.items() if k in user_rights}
    
    num_items = len(NAV_BAR_PAGES) + 1
    cols = st.columns(num_items)
    st.markdown("""
        <style>
        div[data-testid="column"] .stButton>button {
            border-radius: 0px !important; background-color: #800000; color: #FFD700;
            border: 1px solid #FFD700; font-weight: bold; font-size: 15px; letter-spacing: 1.5px;
            height: 4em; width: 100%; transition: all 0.3s ease;
        }
        div[data-testid="column"] .stButton>button:hover { background-color: #A00000; color: #FFFFFF; }
        </style>
    """, unsafe_allow_html=True)
    for i, (key, value) in enumerate(NAV_BAR_PAGES.items()):
        if cols[i].button(value['label'], key=f"nav_{key}"):
            st.session_state.current_page = key
            st.rerun()
    if cols[-1].button("LOGOUT", key="nav_logout"):
        st.session_state.logged_in = False
        st.rerun()
    st.write("---")

# --- APP INIT ---
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "Home Dashboard"

if not st.session_state.logged_in:
    bg_img_base64 = get_base64_of_bin_file(BACKGROUND_PATH)
    st.markdown(f"""
        <style>
        .stApp {{ 
            {"background-image: url('data:image/jpg;base64," + bg_img_base64 + "');" if bg_img_base64 else ""}
            background-size: 75% !important; background-position: center !important;
            background-repeat: no-repeat !important; background-attachment: fixed !important;
            background-color: #bf953f; 
        }}
        label {{ color: #800000 !important; font-weight: 900 !important; font-size: 20px !important; }}
        .stButton>button {{ background-color: #800000; color: #FFD700; border: 1px solid #FFD700; font-weight: bold; }}
        </style>
        """, unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown("<div style='height: 15vh;'></div>", unsafe_allow_html=True)
        st.markdown("<h1 style='font-size: 60px; text-align: center; color: #800000;'>🕉️</h1>", unsafe_allow_html=True)
        st.markdown("<div style='background-color: #800000; color: #FFD700; padding: 15px; text-align: center; font-weight: bold; border: 2px solid #FFD700;'>Temple Management System - Login</div>", unsafe_allow_html=True)
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.button("SIGN IN"):
            success, role, rights = verify_user(username, password)
            if success:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.role = role
                st.session_state.rights = rights
                st.rerun()
            else: st.error("Invalid Login")
    render_footer()
    st.stop()

# --- MAIN APP LAYOUT (GOLD GRADIENT) ---
st.markdown("""
    <style>
    .stApp { background: linear-gradient(135deg, #bf953f 0%, #fcf6ba 25%, #b38728 50%, #fbf5b7 75%, #aa771c 100%); background-attachment: fixed; }
    </style>
    """, unsafe_allow_html=True)

# --- PAGES ---

if st.session_state.current_page == "Home Dashboard":
    page_header(); render_navigation_bar()
    st.title(f"Welcome, {st.session_state.username.title()}")
    df_fam = get_data("families")
    st.metric("Total Devotees", len(df_fam))

elif st.session_state.current_page == "Enroll":
    page_header(); render_navigation_bar()
    st.header("Devotee Enrollment")
    with st.form("enroll_form"):
        hn = st.text_input("Head Name *"); hp = st.text_input("Phone *"); ha = st.text_area("Address")
        if st.form_submit_button("Save Devotee"):
            if hn and hp: run_supabase_insert("families", {"head_name": hn, "phone": hp, "address": ha}); st.success("Saved!")

elif st.session_state.current_page == "Billing":
    page_header(); render_navigation_bar(); st.header("Billing Desk")
    servs = get_data("services")
    if not servs.empty:
        with st.form("bill_form"):
            name = st.text_input("Devotee Name"); s_dict = {r['service_name']: r for _, r in servs.iterrows()}
            sel_s = st.selectbox("Select Service", list(s_dict.keys()))
            if st.form_submit_button("Generate Bill"):
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                run_supabase_insert("transactions", {"guest_name": name, "amount": s_dict[sel_s]['price'], "service_id": s_dict[sel_s]['id'], "date": now, "family_id": 0})
                st.success("Bill Recorded!")

elif st.session_state.current_page == "Expenses":
    page_header(); render_navigation_bar(); st.header("Expenses")
    with st.form("exp_f"):
        en = st.text_input("Title"); ea = st.number_input("Amount")
        if st.form_submit_button("Record"):
            run_supabase_insert("users_expenses", {"expense_name": en, "amount": ea, "payment_date": str(date.today()), "status": "Paid"})
            st.rerun()

elif st.session_state.current_page == "Reports":
    page_header()
    render_navigation_bar()
    st.header("Financial Reports")

    # --- SERVICEWISE FILTER LOGIC ---
    st.subheader("Generate Servicewise Report")
    df_serv = get_data("services")
    
    if not df_serv.empty:
        c_rep1, c_rep2, c_rep3 = st.columns([2, 1, 1])
        # Allow choosing a specific service or All
        service_options = ["All Services"] + df_serv['service_name'].tolist()
        selected_service_name = c_rep1.selectbox("Select Service for Report", service_options)
        
        start_d = c_rep2.date_input("From Date", value=date.today() - timedelta(30))
        end_d = c_rep3.date_input("To Date", value=date.today())

        if st.button("Generate Servicewise Report", type="primary"):
            df_trans = get_data("transactions")
            if not df_trans.empty:
                df_trans['dt'] = pd.to_datetime(df_trans['date']).dt.date
                # Merge to get names
                df_merged = df_trans.merge(df_serv, left_on='service_id', right_on='id', suffixes=('', '_master'))
                
                # Filter by Date
                mask = (df_merged['dt'] >= start_d) & (df_merged['dt'] <= end_d)
                report_df = df_merged[mask]
                
                # Filter by Service if not "All"
                if selected_service_name != "All Services":
                    report_df = report_df[report_df['service_name'] == selected_service_name]
                
                if not report_df.empty:
                    st.success(f"Showing report for: {selected_service_name}")
                    
                    # Summary Metrics
                    total_amt = report_df['amount'].sum()
                    total_count = len(report_df)
                    st.metric(f"Total Collection ({selected_service_name})", f"₹ {total_amt:,.2f}", f"Count: {total_count}")
                    
                    # Detailed Table
                    display_cols = ['dt', 'guest_name', 'service_name', 'amount']
                    final_report = report_df[display_cols].rename(columns={'dt': 'Date', 'guest_name': 'Devotee', 'service_name': 'Service'})
                    st.dataframe(final_report, use_container_width=True, hide_index=True)
                    
                    # Download
                    st.download_button("📥 Download Servicewise Excel", to_excel(final_report), f"Service_Report_{selected_service_name}_{start_d}.xlsx")
                else:
                    st.warning("No transactions found for the selected criteria.")
            else:
                st.info("No transaction data available.")
    else:
        st.error("No services configured in Settings.")

elif st.session_state.current_page == "Settings":
    page_header(); render_navigation_bar(); st.header("Settings")
    with st.form("svc_f"):
        sn = st.text_input("Service Name"); sp = st.number_input("Price")
        if st.form_submit_button("Add"):
            run_supabase_insert("services", {"service_name": sn, "price": sp}); st.rerun()

render_footer()
