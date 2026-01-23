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
        st.error("Supabase credentials not found. Please check st.secrets.")
        st.stop()

supabase = get_supabase_client()

# --- CONFIGURATION & PAGE SETUP ---
st.set_page_config(page_title="Temple Management System", layout="wide", page_icon="🕉️")

ADMIN_ROLE = 'admin'
USER_ROLE = 'user'

# Menu Keys updated with Marriage Bond sections
ALL_MENU_KEYS = [
    "Home Dashboard", "Enroll", "Search", "Billing", 
    "Marriage Bond", "Bond Report", "Expenses", "Reports", "Assets", "Settings"
]

# (Constants like NATCHATHIRAM_OPTIONS, RELATIONSHIP_OPTIONS, etc. remain the same as your original code)
NATCHATHIRAM_OPTIONS = ['Ashwini', 'Bharani', 'Karthigai', 'Rohini', 'Mrigasiram', 'Thiruvathirai', 'Punarpoosam', 'Poosam', 'Ayilyam', 'Magam', 'Poorvam', 'Uthiram', 'Hastham', 'Chithirai', 'Swathi', 'Visakam', 'Anusham', 'Kettai', 'Moolam', 'Pooradam', 'Uthiradam', 'Thiruvonam', 'Avittam', 'Sathayam', 'Poorattathi', 'Uthirattathi', 'Revathi']
RELATIONSHIP_OPTIONS = ['Wife', 'Son', 'Daughter', 'Mother', 'Father', 'Grand Father', 'Grand Mother', 'Guardian', 'Other']
MIN_DATE = date(1940, 1, 1)
TEMPLE_NAME_FULL = "Sree Bhadreshwari Amman Temple Management System"
# ... (Other Constants)

# --- DB WRAPPER FUNCTIONS ---
def run_supabase_insert(table_name, data):
    try: return supabase.table(table_name).insert(data).execute()
    except Exception as e: st.error(f"DB Insert Error: {e}"); return None

def run_supabase_update(table_name, data, row_id):
    try: return supabase.table(table_name).update(data).eq('id', row_id).execute()
    except Exception as e: st.error(f"DB Update Error: {e}"); return None

def run_supabase_delete(table_name, row_id):
    try: return supabase.table(table_name).delete().eq('id', row_id).execute()
    except Exception as e: st.error(f"DB Delete Error: {e}"); return None

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

# (Include your image_to_base64, base64_to_image, etc. from original code)

# --- NAVIGATION ---
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
    # (Insert your login_page() call here)
    st.stop()

# --- MODULES ---

# 1. NEW MODULE: Marriage Bond Entry
if st.session_state.current_page == "Marriage Bond":
    st.header("📜 Marriage Bond Entry (Samaya Vakuppu)")
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
        
        if st.form_submit_button("🚀 Issue & Save Bond Details"):
            if st_name and st_bond_no:
                bond_payload = {
                    "student_name": st_name,
                    "dob": format_date_for_db(st_dob),
                    "issuing_bank": st_bank,
                    "bond_no": st_bond_no,
                    "issuing_date": format_date_for_db(st_issue_date),
                    "maturity_date": format_date_for_db(st_maturity)
                }
                res = run_supabase_insert("marriage_bonds", bond_payload)
                if res: st.success(f"Marriage Bond recorded for {st_name} successfully!"); st.rerun()
            else: st.error("Please fill Student Name and Bond Number.")

# 2. NEW MODULE: Bond Management & Reports
elif st.session_state.current_page == "Bond Report":
    st.header("📋 Marriage Bond Reports")
    bonds_df = get_data("marriage_bonds")
    
    if bonds_df.empty:
        st.info("No Marriage Bonds issued yet.")
    else:
        # Search Filter
        s_term = st.text_input("Search Student or Bond Number")
        if s_term:
            bonds_df = bonds_df[bonds_df['student_name'].str.contains(s_term, case=False, na=False) | 
                                bonds_df['bond_no'].str.contains(s_term, case=False, na=False)]
        
        # Display Table
        disp_df = bonds_df.copy()
        disp_df['dob'] = disp_df['dob'].apply(format_date_for_ui)
        disp_df['issuing_date'] = disp_df['issuing_date'].apply(format_date_for_ui)
        disp_df['maturity_date'] = disp_df['maturity_date'].apply(format_date_for_ui)
        st.dataframe(disp_df, use_container_width=True, hide_index=True)
        
        st.divider()
        st.subheader("⚙️ Manage Bond Records")
        m_tab1, m_tab2 = st.tabs(["Edit Bond", "Delete Bond"])
        
        with m_tab1:
            b_dict = {f"{r['student_name']} (Bond: {r['bond_no']})": r['id'] for _, r in bonds_df.iterrows()}
            b_sel = st.selectbox("Select Record to Edit", list(b_dict.keys()))
            curr_b = bonds_df[bonds_df['id'] == b_dict[b_sel]].iloc[0]
            
            with st.form("edit_bond_form"):
                ec1, ec2 = st.columns(2)
                with ec1:
                    u_name = st.text_input("Student Name", value=curr_b['student_name'])
                    u_dob = st.date_input("DOB", value=safe_date_convert(curr_b['dob']))
                    u_bank = st.text_input("Bank", value=curr_b['issuing_bank'])
                with ec2:
                    u_no = st.text_input("Bond No", value=curr_b['bond_no'])
                    u_issue = st.date_input("Issue Date", value=safe_date_convert(curr_b['issuing_date']))
                    u_maturity = st.date_input("Maturity Date", value=safe_date_convert(curr_b['maturity_date']))
                
                if st.form_submit_button("Update Bond Details"):
                    run_supabase_update("marriage_bonds", {
                        "student_name": u_name, "dob": format_date_for_db(u_dob),
                        "issuing_bank": u_bank, "bond_no": u_no,
                        "issuing_date": format_date_for_db(u_issue),
                        "maturity_date": format_date_for_db(u_maturity)
                    }, curr_b['id'])
                    st.success("Bond Record Updated!"); st.rerun()

        with m_tab2:
            d_b_sel = st.selectbox("Select Record to Permanent Delete", list(b_dict.keys()), key="del_bond")
            if st.button("🚨 Delete Bond Record"):
                run_supabase_delete("marriage_bonds", b_dict[d_b_sel])
                st.warning("Record deleted successfully."); st.rerun()

# --- OTHER MODULES (Enroll, Search, Billing, etc.) ---
# ... (Keep your existing module code here)

render_footer()
