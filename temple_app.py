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
GUEST_FAMILY_ID = 0 

# Default/Fallback Options
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

# --- DB WRAPPER FUNCTIONS (SUPABASE) ---

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

# --- PDF GENERATORS ---

def generate_pdf(receipt_no, devotee_name, devotee_address, service, amount, trans_date, manual_no, book_no):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    y_start = 780 
    try:
        if os.path.exists(LOGO_PATH):
            img = ImageReader(LOGO_PATH)
            c.drawImage(img, 50, y_start - 50, width=50, height=50)
    except: pass
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(300, y_start, TEMPLE_NAME_FULL)
    c.setFont("Helvetica", 10)
    c.drawCentredString(300, y_start - 15, TRUST_DETAILS)
    c.drawCentredString(300, y_start - 30, ADDRESS_LINE_1)
    c.drawCentredString(300, y_start - 45, ADDRESS_LINE_2)
    c.line(50, y_start - 60, 550, y_start - 60)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y_start - 90, f"RECEIPT No: #{receipt_no}")
    c.drawString(400, y_start - 90, f"DATE: {trans_date}")
    c.setFont("Helvetica", 12)
    c.drawString(50, y_start - 140, f"Devotee Name: {devotee_name}")
    c.drawString(50, y_start - 160, f"Address: {devotee_address[:70]}")
    c.drawString(50, y_start - 200, f"Seva / Pooja: {service}")
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, y_start - 240, f"AMOUNT PAID: Rs. {float(amount):,.2f}/-")
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(50, y_start - 340, "Thank you for your offering. May the blessings of Amman be upon you.")
    c.save()
    buffer.seek(0)
    return buffer

def generate_financial_pdf(income_df, expense_df, title, t_inc, t_exp, t_net):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    story = [Paragraph(TEMPLE_NAME_FULL, styles['Title']), Paragraph(title, styles['h3']), Spacer(1, 12)]
    
    summary_data = [["Total Income", f"₹ {t_inc:,.2f}"], ["Total Expenses", f"₹ {t_exp:,.2f}"], ["Net Profit", f"₹ {t_net:,.2f}"]]
    st_table = Table(summary_data, colWidths=[150, 150])
    st_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (0, -1), colors.lightgrey), ('GRID', (0,0), (-1,-1), 0.5, colors.black)]))
    story.append(st_table)
    doc.build(story)
    buffer.seek(0)
    return buffer

# --- INTERFACE COMPONENTS ---

def render_footer():
    st.markdown("""
        <style>
        .footer { position: fixed; left: 0; bottom: 0; width: 100%; background-color: #800000; color: #FFD700; text-align: center; padding: 10px 0; font-size: 14px; font-weight: bold; z-index: 999; border-top: 2px solid #FFD700; }
        </style>
        <div class="footer">Developed By : Sai Dharshini Info Solution</div>
    """, unsafe_allow_html=True)

def render_navigation_bar():
    ALL_PAGES = {"Home Dashboard": "HOME", "Enroll": "ENROLLMENT", "Search": "SEARCH", "Billing": "BILLING", "Expenses": "EXPENSES", "Reports": "REPORTS", "Assets": "ASSETS", "Samayavakuppu": "SAMAYAVAKUPPU", "Settings": "SETTINGS"}
    pages = ALL_PAGES if st.session_state.role == ADMIN_ROLE else {k: v for k, v in ALL_PAGES.items() if k in st.session_state.rights}
    if st.session_state.role == ADMIN_ROLE: pages["Users"] = "USERS"
    
    cols = st.columns(len(pages) + 1)
    for i, (key, label) in enumerate(pages.items()):
        if cols[i].button(label, key=f"nav_{key}", use_container_width=True):
            st.session_state.current_page = key
            st.rerun()
    if cols[-1].button("LOGOUT", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()
    st.write("---")

# --- MAIN APP LOGIC ---

if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "Home Dashboard"

if not st.session_state.logged_in:
    # SIMPLE LOGIN UI
    st.title("🕉️ Staff Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("SIGN IN"):
        s, r, ri = verify_user(u, p)
        if s:
            st.session_state.logged_in, st.session_state.role, st.session_state.rights, st.session_state.username = True, r, ri, u
            st.rerun()
    st.stop()

# --- PAGE ROUTING ---

if st.session_state.current_page == "Home Dashboard":
    st.header(f"Welcome, {st.session_state.username.title()}")
    render_navigation_bar()
    # Metrics logic here...
    st.info("Select a module from the menu above to begin.")

elif st.session_state.current_page == "Enroll":
    render_navigation_bar()
    st.header("Devotee Enrollment")
    # Enrollment logic here...

elif st.session_state.current_page == "Billing":
    render_navigation_bar()
    st.header("Billing Desk")
    
    billing_mode = st.radio("Billing Mode", ["Enrolled Devotee", "Guest Devotee"], horizontal=True)
    
    if billing_mode == "Enrolled Devotee":
        f_data = get_data("families")
        m_data = get_data("members")
        
        if not f_data.empty:
            opts = {}
            for _, f in f_data.iterrows():
                label = f"{f['head_name']} (Head) | 📱 {f['phone']}"
                opts[label] = {"id": f['id'], "name": f['head_name'], "address": f['address'], "wa": f['whatsapp'] if f['whatsapp'] else f['phone']}
            
            for _, m in m_data.iterrows():
                head = f_data[f_data['id'] == m['family_id']]
                h_phone = head['phone'].values[0] if not head.empty else ""
                h_addr = head['address'].values[0] if not head.empty else ""
                m_phone = m['phone'] if m['phone'] else h_phone
                label = f"{m['member_name']} ({m['relationship']}) | 📱 {m_phone}"
                opts[label] = {"id": m['family_id'], "name": m['member_name'], "address": h_addr, "wa": m['whatsapp'] if m['whatsapp'] else m_phone}
            
            sel_k = st.selectbox("Search Devotee by Name or Mobile No", [""] + list(opts.keys()))
            if sel_k:
                d = opts[sel_k]
                st.success(f"Selected: {d['name']}")
                # Billing form logic...
    else:
        guest_name = st.text_input("Guest Name")

elif st.session_state.current_page == "Reports":
    render_navigation_bar()
    st.header("Financial Reports")
    
    report_mode = st.radio("Period:", ["Daily", "Weekly", "Monthly", "Custom"], horizontal=True)
    today = date.today()
    start_d, end_d = today, today # Default
    
    # Date Range Selection Logic (Simplified)
    if report_mode == "Daily": start_d = st.date_input("Date", value=today); end_d = start_d
    elif report_mode == "Custom": 
        c1, c2 = st.columns(2)
        start_d = c1.date_input("Start")
        end_d = c2.date_input("End")
    
    # Fetch Data
    df_trans = get_data("transactions")
    df_exp = get_data("users_expenses")
    
    if not df_trans.empty:
        df_trans['dt'] = pd.to_datetime(df_trans['date']).dt.date
        df_trans = df_trans[(df_trans['dt'] >= start_d) & (df_trans['dt'] <= end_d)]
    if not df_exp.empty:
        df_exp['dt'] = pd.to_datetime(df_exp['payment_date']).dt.date
        df_exp = df_exp[(df_exp['dt'] >= start_d) & (df_exp['dt'] <= end_d)]

    # Metrics
    t_inc = df_trans['amount'].sum() if not df_trans.empty else 0
    t_exp = df_exp['amount'].sum() if not df_exp.empty else 0
    st.divider()
    m1, m2, m3 = st.columns(3)
    m1.metric("Income", f"₹ {t_inc:,.2f}")
    m2.metric("Expenses", f"₹ {t_exp:,.2f}")
    m3.metric("Net Profit", f"₹ {t_inc - t_exp:,.2f}")

    # Build Ledger
    ledger_rows = []
    if not df_trans.empty:
        for _, r in df_trans.iterrows(): 
            ledger_rows.append({"Date": r['dt'], "Description": r['guest_name'] or "Seva", "Income": r['amount'], "Expenses": 0, "Type": "Income"})
    if not df_exp.empty:
        for _, r in df_exp.iterrows(): 
            ledger_rows.append({"Date": r['dt'], "Description": r['expense_name'], "Income": 0, "Expenses": r['amount'], "Type": r['expense_type']})

    if ledger_rows:
        ledger_df = pd.DataFrame(ledger_rows).sort_values("Date")
        
        # CATEGORY-WISE SUMMARY
        st.subheader("📊 Category-Wise Summary")
        cat_sum = ledger_df.groupby('Type').agg({'Income': 'sum', 'Expenses': 'sum'}).reset_index()
        st.table(cat_sum)

        # CHARTS
        st.subheader("📈 Visual Distribution")
        st.bar_chart(cat_sum.set_index('Type')[['Income', 'Expenses']])
        
        # DETAILED TABLE
        st.subheader("📝 Detailed Ledger")
        st.dataframe(ledger_df, use_container_width=True)
    else:
        st.info("No records for this period.")

# ... Other modules (Assets, Samayavakuppu, etc.) would follow similar elif structures

render_footer()
