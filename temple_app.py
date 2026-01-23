import streamlit as st
import pd as pd
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
        st.error("Supabase credentials not found. Check st.secrets.")
        st.stop()

supabase = get_supabase_client()

# --- CONFIGURATION & PAGE SETUP ---
st.set_page_config(page_title="Temple Management System", layout="wide", page_icon="🕉️")

ADMIN_ROLE = 'admin'
USER_ROLE = 'user'

# Menu Keys updated with Marriage Bond sections
ALL_MENU_KEYS = [
    "Home Dashboard", "Enroll", "Search", "Billing", 
    "Marriage Bond", "Expenses", "Reports", "Assets", "Settings"
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

# --- DB WRAPPER FUNCTIONS ---
def run_supabase_insert(table_name, data):
    try: return supabase.table(table_name).insert(data).execute()
    except Exception as e: st.error(f"Insert Error: {e}"); return None

def run_supabase_update(table_name, data, row_id):
    try: return supabase.table(table_name).update(data).eq('id', row_id).execute()
    except Exception as e: st.error(f"Update Error: {e}"); return None

def run_supabase_delete(table_name, row_id):
    try: return supabase.table(table_name).delete().eq('id', row_id).execute()
    except Exception as e: st.error(f"Delete Error: {e}"); return None

def get_data(table_name, select="*"):
    try:
        res = supabase.table(table_name).select(select).execute()
        return pd.DataFrame(res.data)
    except: return pd.DataFrame()

# --- UTILITY FUNCTIONS ---
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

# --- NAVIGATION & LOGIN ---
def render_navigation_bar():
    ALL_PAGES = {
        "Home Dashboard": {"label": "HOME"}, 
        "Enroll": {"label": "ENROLLMENT"},
        "Search": {"label": "SEARCH"}, 
        "Billing": {"label": "BILLING"},
        "Marriage Bond": {"label": "MARRIAGE BOND"}, # Integrated
        "Expenses": {"label": "EXPENSES"}, 
        "Reports": {"label": "REPORTS"},
        "Assets": {"label": "ASSETS"}, 
        "Settings": {"label": "SETTINGS"},
    }
    
    user_rights = st.session_state.get('rights', ["Home Dashboard"])
    if st.session_state.role == ADMIN_ROLE:
        NAV_BAR_PAGES = ALL_PAGES.copy()
        NAV_BAR_PAGES["Users"] = {"label": "USERS"}
    else:
        NAV_BAR_PAGES = {k: v for k, v in ALL_PAGES.items() if k in user_rights}
    
    cols = st.columns(len(NAV_BAR_PAGES) + 1)
    st.markdown("""<style>div[data-testid="column"] .stButton>button { border-radius: 0px !important; background-color: #800000; color: #FFD700; border: 1px solid #FFD700; font-weight: bold; font-size: 14px; height: 3.5em; width: 100%; }</style>""", unsafe_allow_html=True)
    
    for i, (key, value) in enumerate(NAV_BAR_PAGES.items()):
        if cols[i].button(value['label'], key=f"nav_{key}", use_container_width=True):
            st.session_state.current_page = key
            st.rerun()
    if cols[-1].button("LOGOUT", key="nav_logout", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()
    st.write("---")

# --- LOGIN FLOW (Simplified for logic) ---
# (Assume existing verify_user and login_page functions from your previous turn are present)

# --- MODULES ---

if st.session_state.current_page == "Marriage Bond":
    st.header("📜 Marriage Bond Management (Samaya Vakuppu)")
    
    m_tab1, m_tab2 = st.tabs(["📝 Issue New Bond", "📋 Bond Reports & Management"])
    
    with m_tab1:
        st.subheader("Student Marriage Bond Registration")
        with st.form("bond_entry_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                st_name = st.text_input("Name of the Student *")
                st_dob = st.date_input("Date of Birth", value=None, min_value=MIN_DATE)
                st_bank = st.text_input("Bond Issuing Bank")
            with c2:
                st_bond_no = st.text_input("Bond Number *")
                st_issue_date = st.date_input("Issuing Date", value=date.today())
                st_maturity = st.date_input("Maturity Date", value=None)
            
            if st.form_submit_button("🚀 Save Bond Details"):
                if st_name and st_bond_no:
                    payload = {
                        "student_name": st_name,
                        "dob": format_date_for_db(st_dob),
                        "issuing_bank": st_bank,
                        "bond_no": st_bond_no,
                        "issuing_date": format_date_for_db(st_issue_date),
                        "maturity_date": format_date_for_db(st_maturity)
                    }
                    run_supabase_insert("marriage_bonds", payload)
                    st.success(f"Marriage Bond recorded for {st_name}!"); st.rerun()
                else: st.error("Student Name and Bond Number are required.")

    with m_tab2:
        st.subheader("Issued Marriage Bonds Search")
        bonds_df = get_data("marriage_bonds")
        if bonds_df.empty:
            st.info("No bonds recorded.")
        else:
            s_term = st.text_input("Search Student or Bond No.")
            if s_term:
                bonds_df = bonds_df[bonds_df['student_name'].str.contains(s_term, case=False, na=False) | 
                                    bonds_df['bond_no'].str.contains(s_term, case=False, na=False)]
            
            # Formatted display
            disp_df = bonds_df.copy()
            for col in ['dob', 'issuing_date', 'maturity_date']:
                disp_df[col] = disp_df[col].apply(format_date_for_ui)
            
            st.dataframe(disp_df.drop(columns=['created_at']), use_container_width=True, hide_index=True)
            
            st.divider()
            st.markdown("### ⚙️ Edit / Delete Bond Records")
            b_dict = {f"{r['student_name']} (Bond: {r['bond_no']})": r['id'] for _, r in bonds_df.iterrows()}
            b_sel = st.selectbox("Select Bond to Modify", list(b_dict.keys()))
            curr_b = bonds_df[bonds_df['id'] == b_dict[b_sel]].iloc[0]
            
            with st.form("edit_bond_form"):
                ec1, ec2 = st.columns(2)
                with ec1:
                    u_name = st.text_input("Student Name", value=curr_b['student_name'])
                    u_bank = st.text_input("Bank", value=curr_b['issuing_bank'])
                with ec2:
                    u_no = st.text_input("Bond No", value=curr_b['bond_no'])
                    u_maturity = st.date_input("Maturity Date", value=safe_date_convert(curr_b['maturity_date']))
                
                edit_col, del_col = st.columns([1, 4])
                if edit_col.form_submit_button("Update"):
                    run_supabase_update("marriage_bonds", {
                        "student_name": u_name, "issuing_bank": u_bank, "bond_no": u_no, 
                        "maturity_date": format_date_for_db(u_maturity)
                    }, curr_b['id'])
                    st.success("Updated!"); st.rerun()
                
                if del_col.form_submit_button("🚨 Permanent Delete"):
                    run_supabase_delete("marriage_bonds", curr_b['id'])
                    st.warning("Deleted record."); st.rerun()

# --- (Other modules: Home, Enroll, Search, Billing, Expenses, Reports, etc. follow here) ---
# Ensure you maintain the same current_page checking logic for those modules.
