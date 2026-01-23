import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date
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

# --- CONFIGURATION & PAGE SETUP ---
# Must be the very first Streamlit command
st.set_page_config(page_title="Sree Bhadreshwari Amman Temple Management", layout="wide", page_icon="🕉️")

ADMIN_ROLE = 'admin'
USER_ROLE = 'user'

# Options
RELATIONSHIP_OPTIONS = ['Wife', 'Son', 'Daughter', 'Mother', 'Father', 'Grand Father', 'Grand Mother', 'Guardian', 'Other']
NATCHATHIRAM_OPTIONS = ['Ashwini', 'Bharani', 'Karthigai', 'Rohini', 'Mrigasiram', 'Thiruvathirai', 'Punarpoosam', 'Poosam', 'Ayilyam', 'Magam', 'Poorvam', 'Uthiram', 'Hastham', 'Chithirai', 'Swathi', 'Visakam', 'Anusham', 'Kettai', 'Moolam', 'Pooradam', 'Uthiradam', 'Thiruvonam', 'Avittam', 'Sathayam', 'Poorattathi', 'Uthirattathi', 'Revathi']
EXPENSE_TYPES = ['Pooja Items', 'Maintenance/Repairs', 'Salary/Dakshina', 'Electricity/Water', 'Annadanam/Food', 'Construction', 'Festivals', 'Administrative', 'Other']
PAYMENT_METHODS = ['Cash', 'UPI / GPay', 'Bank Transfer', 'Cheque', 'Card']
PAYMENT_STATUS = ['Paid', 'Pending', 'Partial']

# Date constraints
MIN_DATE = date(1940, 1, 1)
MAX_DATE = date(2040, 12, 31)

# TEMPLE DETAILS
TEMPLE_NAME_FULL = "Sree Bhadreshwari Amman Temple"
TRUST_DETAILS = "Samrakshana Seva Trust 174/2004"
ADDRESS_LINE_1 = "Kanjampuram"
ADDRESS_LINE_2 = "Kanniyakumari Dist., Tamil Nadu - 629154"
LOGO_PATH = "amman.jpg" 
BACKGROUND_PATH = "background.jpg"

# --- UTILITY FUNCTIONS ---

def get_whatsapp_link(phone, devotee_name, service_name, amount, receipt_id):
    """Formats a WhatsApp API link with encoded temple receipt details."""
    if phone:
        clean_phone = "".join(filter(str.isdigit, str(phone)))
        if len(clean_phone) == 10: clean_phone = f"91{clean_phone}"
        message = (
            f"🙏 *Greetings from {TEMPLE_NAME_FULL}*\n\n"
            f"Namaste {devotee_name},\n"
            f"We have received your offering for *{service_name}*.\n"
            f"Receipt No: #{receipt_id}\n"
            f"Amount Paid: ₹{amount:,.2f}\n\n"
            f"May the blessings of Sree Bhadreshwari Amman be with you. ✨"
        )
        return f"https://wa.me/{clean_phone}?text={urllib.parse.quote(message)}"
    return None

def image_to_base64(image_file):
    if image_file: return base64.b64encode(image_file.getvalue()).decode()
    return ""

def base64_to_image(base64_str):
    if base64_str:
        try: return io.BytesIO(base64.b64decode(base64_str))
        except: return None
    return None

def init_db():
    conn = sqlite3.connect('temple_data.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS families (id INTEGER PRIMARY KEY AUTOINCREMENT, head_name TEXT, dob DATE, wedding_date DATE, natchathiram TEXT, address TEXT, phone TEXT, whatsapp TEXT, photo TEXT, yearly_pooja_date DATE)''')
    c.execute('''CREATE TABLE IF NOT EXISTS members (id INTEGER PRIMARY KEY AUTOINCREMENT, family_id INTEGER, member_name TEXT, relationship TEXT, dob DATE, wedding_date DATE, natchathiram TEXT, yearly_pooja_date DATE, phone TEXT, whatsapp TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS services (id INTEGER PRIMARY KEY AUTOINCREMENT, service_name TEXT, price REAL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS transactions (id INTEGER PRIMARY KEY AUTOINCREMENT, family_id INTEGER, service_id INTEGER, amount REAL, date DATETIME, manual_bill_no TEXT, bill_book_no TEXT, guest_name TEXT, guest_address TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS users_expenses (id INTEGER PRIMARY KEY AUTOINCREMENT, expense_name TEXT, expense_type TEXT, amount REAL, payment_date DATE, status TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password_hash TEXT, role TEXT)''')
    # Marriage Bond Table
    c.execute('''CREATE TABLE IF NOT EXISTS marriage_bonds (id INTEGER PRIMARY KEY AUTOINCREMENT, student_name TEXT, dob DATE, issuing_bank TEXT, bond_no TEXT, issuing_date DATE, maturity_date DATE)''')
    
    c.execute("SELECT count(*) FROM users WHERE username='admin'")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)", ('admin', hashlib.sha256('admin123'.encode()).hexdigest(), ADMIN_ROLE))
    conn.commit()
    conn.close()

def run_query(query, params=()):
    conn = sqlite3.connect('temple_data.db'); c = conn.cursor()
    c.execute(query, params); conn.commit(); conn.close()

def get_data(query, params=()):
    conn = sqlite3.connect('temple_data.db')
    df = pd.read_sql(query, conn, params=params); conn.close()
    return df

def generate_pdf(receipt_no, devotee_name, devotee_address, service, amount, trans_date, manual_no, book_no):
    buffer = io.BytesIO(); c = canvas.Canvas(buffer, pagesize=letter); y_pos = 780 
    c.setFont("Helvetica-Bold", 16); c.drawCentredString(300, y_pos, TEMPLE_NAME_FULL)
    c.setFont("Helvetica", 10); c.drawCentredString(300, y_pos-15, TRUST_DETAILS)
    c.line(50, y_pos-50, 550, y_pos-50)
    c.drawString(50, y_pos-80, f"Receipt: #{receipt_no}"); c.drawString(400, y_pos-80, f"Date: {trans_date}")
    c.drawString(50, y_pos-110, f"Devotee: {devotee_name}"); c.drawString(50, y_pos-140, f"Service: {service}")
    c.setFont("Helvetica-Bold", 14); c.drawString(50, y_pos-180, f"Total Paid: Rs. {amount:,.2f}")
    c.save(); buffer.seek(0); return buffer

# --- NAVIGATION & LOGIN ---

def render_navigation_bar():
    pages = {
        "Home": "🏠 HOME", "Enroll": "📝 ENROLL", "Search": "🔍 SEARCH", 
        "Billing": "💵 BILLING", "Bond Entry": "📜 BOND ENTRY", 
        "Bond Reports": "📊 BOND REPORTS", "Expenses": "💸 EXPENSES", "Financials": "📑 REPORTS"
    }
    cols = st.columns(len(pages) + 1)
    for i, (key, label) in enumerate(pages.items()):
        if cols[i].button(label, use_container_width=True):
            st.session_state.current_page = key
            st.rerun()
    if cols[-1].button("🚪 LOGOUT", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

init_db()
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "Home"

if not st.session_state.logged_in:
    st.title("Temple Management Login")
    u, p = st.text_input("Username"), st.text_input("Password", type="password")
    if st.button("Login"):
        if u == 'admin' and p == 'admin123':
            st.session_state.logged_in = True; st.session_state.username = u; st.rerun()
    st.stop()

render_navigation_bar()

# --- APP MODULES ---

if st.session_state.current_page == "Home":
    st.title(f"Welcome to Sree Bhadreshwari Amman Temple Portal")
    st.markdown("---")
    st.info("Please select a menu above to process devotee registrations, bills, or bonds.")

elif st.session_state.current_page == "Billing":
    st.header("Billing Desk")
    mode = st.radio("Devotee Type", ["Enrolled", "Guest"], horizontal=True)
    
    col_in, col_h = st.columns([2, 1])
    with col_in:
        if mode == "Enrolled":
            fams = get_data("SELECT id, head_name, phone, address FROM families")
            if not fams.empty:
                fam_dict = {f"{r['head_name']} ({r['phone']})": r for _, r in fams.iterrows()}
                sel = st.selectbox("Select Devotee", list(fam_dict.keys()))
                devotee = fam_dict[sel]
                s_name, s_phone, s_addr, s_id = devotee['head_name'], devotee['phone'], devotee['address'], devotee['id']
            else: st.warning("No devotees enrolled."); st.stop()
        else:
            s_name, s_phone, s_addr, s_id = st.text_input("Guest Name"), st.text_input("WhatsApp No"), st.text_area("Address"), 0

        servs = get_data("SELECT * FROM services")
        if servs.empty: st.error("Please add services in Settings."); st.stop()
        s_dict = {r['service_name']: r for _, r in servs.iterrows()}
        serv = s_dict[st.selectbox("Select Pooja", list(s_dict.keys()))]
        
        if st.button("Generate Bill & Send Notification"):
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            run_query("INSERT INTO transactions (family_id, service_id, amount, date, guest_name, guest_address) VALUES (?,?,?,?,?,?)", (s_id, serv['id'], serv['price'], now, s_name if s_id==0 else "", s_addr if s_id==0 else ""))
            last_id = get_data("SELECT max(id) as id FROM transactions").iloc[0]['id']
            pdf = generate_pdf(last_id, s_name, s_addr, serv['service_name'], serv['price'], now, "", "")
            
            st.success(f"Receipt #{last_id} Created!")
            c1, c2 = st.columns(2)
            with c1: st.download_button("📥 Download PDF", pdf, f"Bill_{last_id}.pdf", use_container_width=True)
            with c2: 
                if s_phone: st.link_button("📲 Send WhatsApp", get_whatsapp_link(s_phone, s_name, serv['service_name'], serv['price'], last_id), use_container_width=True)

elif st.session_state.current_page == "Bond Entry":
    st.header("Marriage Bond Entry - Samaya Vakuppu")
    with st.form("bond_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            name = st.text_input("Student Name *")
            dob = st.date_input("Date of Birth", value=None)
            bank = st.text_input("Bond Issuing Bank")
        with c2:
            b_no = st.text_input("Bond No *")
            i_date = st.date_input("Issuing Date", value=date.today())
            m_date = st.date_input("Maturity Date")
        if st.form_submit_button("Save Bond Details"):
            if name and b_no:
                run_query("INSERT INTO marriage_bonds (student_name, dob, issuing_bank, bond_no, issuing_date, maturity_date) VALUES (?,?,?,?,?,?)", (name, dob, bank, b_no, i_date, m_date))
                st.success(f"Bond for {name} saved!")
            else: st.error("Name and Bond No are mandatory.")

elif st.session_state.current_page == "Bond Reports":
    st.header("Issued Bonds Report")
    search = st.text_input("Search Student Name or Bond Number")
    q = "SELECT * FROM marriage_bonds"
    if search: q += f" WHERE student_name LIKE '%{search}%' OR bond_no LIKE '%{search}%'"
    df = get_data(q)
    
    for _, row in df.iterrows():
        with st.expander(f"📜 {row['student_name']} (Bond: {row['bond_no']})"):
            with st.form(f"eb_{row['id']}"):
                u_name = st.text_input("Student Name", value=row['student_name'])
                u_bank = st.text_input("Bank", value=row['issuing_bank'])
                u_no = st.text_input("Bond Number", value=row['bond_no'])
                cb1, cb2 = st.columns([1, 5])
                if cb1.form_submit_button("Update"):
                    run_query("UPDATE marriage_bonds SET student_name=?, issuing_bank=?, bond_no=? WHERE id=?", (u_name, u_bank, u_no, row['id']))
                    st.rerun()
                if cb2.form_submit_button("🗑️ Delete"):
                    run_query("DELETE FROM marriage_bonds WHERE id=?", (row['id'],))
                    st.rerun()

# --- FOOTER ---
st.markdown("""<style>.footer { position: fixed; left: 0; bottom: 0; width: 100%; background-color: #800000; color: #FFD700; text-align: center; padding: 10px; font-weight: bold; border-top: 2px solid #FFD700; }</style><div class="footer">Developed By : Sai Dharshini Info Solution</div>""", unsafe_allow_html=True)
