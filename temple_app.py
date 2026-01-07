import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date
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

# --- CONFIGURATION & PAGE SETUP ---
# Streamlit set_page_config MUST be the first command
st.set_page_config(page_title="Temple Management System", layout="wide", page_icon="🕉️")

ADMIN_ROLE = 'admin'
USER_ROLE = 'user'
GUEST_FAMILY_ID = 0 

# Dropdown Options
RELATIONSHIP_OPTIONS = ['Wife', 'Son', 'Daughter', 'Mother', 'Father', 'Grand Father', 'Grand Mother', 'Guardian', 'Other']
NATCHATHIRAM_OPTIONS = ['Ashwini', 'Bharani', 'Karthigai', 'Rohini', 'Mrigasiram', 'Thiruvathirai', 'Punarpoosam', 'Poosam', 'Ayilyam', 'Magam', 'Poorvam', 'Uthiram', 'Hastham', 'Chithirai', 'Swathi', 'Visakam', 'Anusham', 'Kettai', 'Moolam', 'Pooradam', 'Uthiradam', 'Thiruvonam', 'Avittam', 'Sathayam', 'Poorattathi', 'Uthirattathi', 'Revathi']
EXPENSE_TYPES = ['Pooja Items', 'Maintenance/Repairs', 'Salary/Dakshina', 'Electricity/Water', 'Annadanam/Food', 'Construction', 'Festivals', 'Administrative', 'Other']
PAYMENT_METHODS = ['Cash', 'UPI / GPay', 'Bank Transfer', 'Cheque', 'Card']
PAYMENT_STATUS = ['Paid', 'Pending', 'Partial']

# Date constraints
MIN_DATE = date(1940, 1, 1)
MAX_DATE = date(2040, 12, 31)

# TEMPLE DETAILS
TEMPLE_NAME_FULL = "Sree Bhadreshwari Amman Temple Management System"
TRUST_DETAILS = "Samrakshana Seva Trust 174/2004"
ADDRESS_LINE_1 = "Kanjampuram"
ADDRESS_LINE_2 = "Kanniyakumari Dist., Tamil Nadu - 629154"
LOGO_PATH = "amman.jpg" 
BACKGROUND_PATH = "background.jpg"

# --- UTILITY FUNCTIONS ---

def get_base64_of_bin_file(bin_file):
    """Converts a local binary file (like an image) to a base64 string."""
    try:
        if os.path.exists(bin_file):
            with open(bin_file, 'rb') as f:
                data = f.read()
            return base64.b64encode(data).decode()
    except:
        pass
    return None

def image_to_base64(image_file):
    """Converts uploaded image to base64 for database storage."""
    if image_file is not None:
        return base64.b64encode(image_file.getvalue()).decode()
    return ""

def base64_to_image(base64_str):
    """Converts base64 string back to bytes for display."""
    if base64_str:
        try:
            return io.BytesIO(base64.b64decode(base64_str))
        except:
            return None
    return None

def to_excel(df):
    """Converts DataFrame to Excel (XLSX). Requires: pip install xlsxwriter"""
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Income_Report')
        return output.getvalue()
    except ModuleNotFoundError:
        st.error("Error: 'xlsxwriter' module not found. Please install it using 'pip install xlsxwriter'.")
        return None

def generate_income_pdf(df, title, total_income):
    """Generates a professional income report PDF."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=50, bottomMargin=30)
    styles = getSampleStyleSheet()
    story = []

    # Header
    story.append(Paragraph(TEMPLE_NAME_FULL, styles['Title']))
    story.append(Paragraph(title, styles['h3']))
    story.append(Spacer(1, 12))
    
    # Summary
    story.append(Paragraph(f"<b>Total Income:</b> ₹ {total_income:,.2f}", styles['Normal']))
    story.append(Spacer(1, 12))

    # Table formatting
    df_pdf = df.copy()
    if 'Amount' in df_pdf.columns:
        df_pdf['Amount'] = df_pdf['Amount'].apply(lambda x: f"₹ {x:,.2f}")
    
    data = [df_pdf.columns.tolist()] + df_pdf.values.tolist()

    t = Table(data, hAlign='LEFT')
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#800000')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.gold),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
    ]))
    
    story.append(t)
    doc.build(story)
    buffer.seek(0)
    return buffer

def init_db():
    """Initializes the SQLite database."""
    conn = sqlite3.connect('temple_data.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS families (id INTEGER PRIMARY KEY AUTOINCREMENT, head_name TEXT, dob DATE, wedding_date DATE, natchathiram TEXT, address TEXT, phone TEXT, whatsapp TEXT, photo TEXT, yearly_pooja_date DATE)''')
    c.execute('''CREATE TABLE IF NOT EXISTS members (id INTEGER PRIMARY KEY AUTOINCREMENT, family_id INTEGER, member_name TEXT, relationship TEXT, dob DATE, wedding_date DATE, natchathiram TEXT, yearly_pooja_date DATE, phone TEXT, whatsapp TEXT, FOREIGN KEY (family_id) REFERENCES families (id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS services (id INTEGER PRIMARY KEY AUTOINCREMENT, service_name TEXT, price REAL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS transactions (id INTEGER PRIMARY KEY AUTOINCREMENT, family_id INTEGER, service_id INTEGER, amount REAL, date DATETIME, manual_bill_no TEXT, bill_book_no TEXT, guest_name TEXT, guest_address TEXT, guest_whatsapp TEXT, FOREIGN KEY (family_id) REFERENCES families (id), FOREIGN KEY (service_id) REFERENCES services (id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS assets (id INTEGER PRIMARY KEY AUTOINCREMENT, asset_name TEXT, description TEXT, value REAL, quantity INTEGER)''')
    c.execute('''CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password_hash TEXT, role TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS users_expenses (id INTEGER PRIMARY KEY AUTOINCREMENT, expense_name TEXT, description TEXT, expense_type TEXT, amount REAL, payment_method TEXT, status TEXT, payment_date DATE, voucher_no TEXT)''')

    # Migrations for families
    try: c.execute("SELECT manual_bill_no FROM transactions LIMIT 1")
    except sqlite3.OperationalError: c.execute("ALTER TABLE transactions ADD COLUMN manual_bill_no TEXT")
    try: c.execute("SELECT bill_book_no FROM transactions LIMIT 1")
    except sqlite3.OperationalError: c.execute("ALTER TABLE transactions ADD COLUMN bill_book_no TEXT")
    try: c.execute("SELECT photo FROM families LIMIT 1")
    except sqlite3.OperationalError: c.execute("ALTER TABLE families ADD COLUMN photo TEXT")
    try: c.execute("SELECT guest_name FROM transactions LIMIT 1")
    except sqlite3.OperationalError: c.execute("ALTER TABLE transactions ADD COLUMN guest_name TEXT")
    try: c.execute("SELECT guest_address FROM transactions LIMIT 1")
    except sqlite3.OperationalError: c.execute("ALTER TABLE transactions ADD COLUMN guest_address TEXT")
    try: c.execute("SELECT yearly_pooja_date FROM families LIMIT 1")
    except sqlite3.OperationalError: c.execute("ALTER TABLE families ADD COLUMN yearly_pooja_date DATE")
    
    # Migrations for WhatsApp and Guests
    try: c.execute("SELECT guest_whatsapp FROM transactions LIMIT 1")
    except sqlite3.OperationalError: c.execute("ALTER TABLE transactions ADD COLUMN guest_whatsapp TEXT")

    # Migrations for members
    try: c.execute("SELECT phone FROM members LIMIT 1")
    except sqlite3.OperationalError: c.execute("ALTER TABLE members ADD COLUMN phone TEXT")
    try: c.execute("SELECT whatsapp FROM members LIMIT 1")
    except sqlite3.OperationalError: c.execute("ALTER TABLE members ADD COLUMN whatsapp TEXT")
    
    c.execute("SELECT count(*) FROM services")
    if c.fetchone()[0] == 0:
        default_services = [('Archanai', 20), ('Neeranchanam', 50), ('Pala Archanai', 100)]
        c.executemany("INSERT INTO services (service_name, price) VALUES (?, ?)", default_services)
    c.execute("SELECT count(*) FROM users WHERE username='admin'")
    if c.fetchone()[0] == 0:
        admin_hash = hashlib.sha256('admin123'.encode()).hexdigest()
        c.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)", ('admin', admin_hash, ADMIN_ROLE))
    c.execute("SELECT count(*) FROM users WHERE username='user'")
    if c.fetchone()[0] == 0:
        user_hash = hashlib.sha256('user123'.encode()).hexdigest()
        c.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)", ('user', user_hash, USER_ROLE))
    conn.commit()
    conn.close()

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def verify_user(username, password):
    conn = sqlite3.connect('temple_data.db')
    c = conn.cursor()
    c.execute("SELECT password_hash, role FROM users WHERE username=?", (username,))
    result = c.fetchone()
    conn.close()
    if result:
        stored_hash, role = result
        if stored_hash == hash_password(password): return True, role
    return False, None

def run_query(query, params=()):
    conn = sqlite3.connect('temple_data.db')
    c = conn.cursor()
    c.execute(query, params)
    conn.commit()
    conn.close()

def get_data(query, params=()):
    conn = sqlite3.connect('temple_data.db')
    df = pd.read_sql(query, conn, params=params)
    conn.close()
    return df

# --- FOOTER COMPONENT ---
def render_footer():
    """Renders a fixed themed footer at the bottom of the page."""
    st.markdown("""
        <style>
        .footer {
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            background-color: #800000;
            color: #FFD700;
            text-align: center;
            padding: 10px 0;
            font-size: 14px;
            font-weight: bold;
            letter-spacing: 1px;
            z-index: 999;
            border-top: 2px solid #FFD700;
        }
        .main .block-container {
            padding-bottom: 60px;
        }
        </style>
        <div class="footer">
            Developed By : Sai Dharshini Info Solution
        </div>
    """, unsafe_allow_html=True)

# --- PDF GENERATOR ---
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
    
    line_y = y_start - 60
    c.line(50, line_y, 550, line_y)

    detail_y_start = line_y - 30
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, detail_y_start, f"RECEIPT No: #{receipt_no}")
    c.drawString(400, detail_y_start, f"DATE: {trans_date}")
    
    c.setFont("Helvetica", 10)
    c.drawString(50, detail_y_start - 20, f"Manual Bill No: {manual_no if manual_no else 'N/A'}")
    c.drawString(250, detail_y_start - 20, f"Bill Book No: {book_no if book_no else 'N/A'}")

    c.setFont("Helvetica", 12)
    c.drawString(50, detail_y_start - 50, f"Devotee Name: {devotee_name}")
    
    addr_text = devotee_address if devotee_address else "N/A"
    c.setFont("Helvetica", 10)
    c.drawString(50, detail_y_start - 70, f"Address: {addr_text[:70]}")

    pooja_y = detail_y_start - 100
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, pooja_y, f"Seva / Pooja: {service}")
    
    amount_y = pooja_y - 40
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, amount_y, "AMOUNT PAID:")
    c.setFillColorRGB(0.5, 0, 0) 
    c.drawString(180, amount_y, f"Rs. {amount:,.2f}/-")
    c.setFillColorRGB(0, 0, 0) 

    footer_y = amount_y - 100
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(50, footer_y, "Thank you for your offering. May the blessings of Sree Bhadreshwari Amman be upon you.")
    c.setFont("Helvetica", 10)
    c.drawString(400, footer_y - 50, "Authorized Signature")
    
    c.save()
    buffer.seek(0)
    return buffer

# --- LOGIN PAGE UI ---
def login_page():
    bg_img_base64 = get_base64_of_bin_file(BACKGROUND_PATH)
    
    # Custom CSS for Login Page
    bg_style = f"""
        <style>
        .stApp {{ 
            {"background-image: url('data:image/jpg;base64," + bg_img_base64 + "');" if bg_img_base64 else ""}
            background-size: 75% !important; 
            background-position: center !important;
            background-repeat: no-repeat !important;
            background-attachment: fixed !important;
            background-color: #bf953f; 
        }}
        /* INCREASED FONT SIZE AND BOLD FOR USERNAME/PASSWORD LABELS */
        label {{ 
            color: #800000 !important; 
            font-weight: 900 !important; 
            font-size: 20px !important; 
            text-align: left !important; 
        }}
        .stButton>button {{ 
            width: 100%; 
            background-color: #800000; 
            color: #FFD700; 
            padding: 10px; 
            border-radius: 0px !important; 
            border: 1px solid #FFD700; 
            font-size: 16px; 
            font-weight: bold; 
            transition: all 0.3s; 
        }}
        .stButton>button:hover {{ background-color: #4D0000; border-color: #FFFFFF; transform: translateY(-2px); }}
        </style>
        """
    st.markdown(bg_style, unsafe_allow_html=True)

    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown("<div style='height: 15vh;'></div>", unsafe_allow_html=True)
        st.markdown("<h1 style='font-size: 60px; text-align: center; color: #800000;'>🕉️</h1>", unsafe_allow_html=True)
        st.markdown("<p style='color: #800000; font-weight: bold; text-align: center;'>Amme Narayana... Devi Narayana...</p>", unsafe_allow_html=True)
        
        # Title Box
        st.markdown("""
            <div style='background-color: #800000; color: #FFD700; padding: 15px; text-align: center; font-weight: bold; font-size: 24px; border: 2px solid #FFD700; margin-bottom: 20px;'>
                Temple Management System
            </div>
        """, unsafe_allow_html=True)
        
        # STAFF LOGIN BOX - MAROON WITH BOLD GOLD TEXT
        st.markdown("""
            <div style='background-color: #800000; color: #FFD700; padding: 10px; text-align: center; font-weight: 900; font-size: 22px; border: 1px solid #FFD700; margin-bottom: 15px;'>
                Staff Login
            </div>
        """, unsafe_allow_html=True)
        
        username = st.text_input("Username", key="login_user")
        password = st.text_input("Password", type="password", key="login_pass")
        if st.button("SIGN IN", key="login_button"):
            success, role = verify_user(username, password)
            if success:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.role = role
                st.rerun()
            else:
                st.error("Invalid Username or Password.")

# --- HEADER COMPONENT ---
def page_header():
    st.markdown("<style>.stToolbar {visibility: hidden;}</style>", unsafe_allow_html=True)
    col_img, col_title = st.columns([1, 8])
    with col_img:
        try: st.image(LOGO_PATH, width=80)
        except: st.markdown("<h3>🕉️</h3>", unsafe_allow_html=True)
    with col_title:
        st.markdown(f"<h1 style='color: #800000; font-size: 32px; padding-top: 10px; border-bottom: 2px solid #b38728;'>{TEMPLE_NAME_FULL}</h1>", unsafe_allow_html=True)
    st.markdown("---")

# --- NAVIGATION BAR ---
def render_navigation_bar():
    NAV_BAR_PAGES = {
        "Home Dashboard": {"label": "HOME"},
        "Enroll": {"label": "ENROLLMENT"},
        "Search": {"label": "SEARCH"},
        "Billing": {"label": "BILLING"},
        "Expenses": {"label": "EXPENSES"},
        "Reports": {"label": "REPORTS"},
        "Assets": {"label": "ASSETS"},
        "Settings": {"label": "SETTINGS"},
    }
    if st.session_state.role == ADMIN_ROLE:
        NAV_BAR_PAGES["Users"] = {"label": "USERS"}
        
    num_items = len(NAV_BAR_PAGES) + 1
    cols = st.columns(num_items)
    
    st.markdown("""
        <style>
        div[data-testid="column"] .stButton>button {
            border-radius: 0px !important; 
            background-color: #800000;
            color: #FFD700;
            border: 1px solid #FFD700;
            font-weight: bold;
            font-size: 15px;
            letter-spacing: 1.5px;
            height: 4em;
            width: 100%;
            margin-top: -10px;
            transition: all 0.3s ease;
            text-transform: uppercase;
            box-shadow: 2px 2px 5px rgba(0,0,0,0.2);
        }
        div[data-testid="column"] .stButton>button:hover {
            background-color: #A00000;
            color: #FFFFFF;
            border-color: #FFFFFF;
            box-shadow: 0 6px 15px rgba(0,0,0,0.3);
            transform: translateY(-2px);
        }
        div[data-testid="column"] .stButton>button:active {
            background-color: #FFD700;
            color: #800000;
        }
        </style>
    """, unsafe_allow_html=True)

    for i, (key, value) in enumerate(NAV_BAR_PAGES.items()):
        if cols[i].button(value['label'], key=f"nav_{key}", use_container_width=True):
            st.session_state.current_page = key
            st.rerun()
            
    if cols[-1].button("LOGOUT", key="nav_logout", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()
    st.write("---")

# --- DATABASE HELPER ---
def save_family_head(head_name, dob, wedding_date, n_star, address, phone, whatsapp, photo_str, pooja_date):
    if head_name and phone:
        conn = sqlite3.connect('temple_data.db')
        c = conn.cursor()
        c.execute("INSERT INTO families (head_name, dob, wedding_date, natchathiram, address, phone, whatsapp, photo, yearly_pooja_date) VALUES (?,?,?,?,?,?,?,?,?)", 
                  (head_name, dob, wedding_date, n_star, address, phone, whatsapp, photo_str, pooja_date))
        conn.commit()
        last_id = c.lastrowid
        conn.close()
        st.session_state.new_family_id = last_id
        st.session_state.new_family_head = head_name
        st.success(f"Family Head {head_name} saved successfully!")
        st.rerun()

# --- TICKER HELPER ---
def render_news_ticker():
    """Fetches today's poojas, birthdays, and anniversaries to display in a continuous marquee."""
    today = date.today()
    today_md = today.strftime('%m-%d')
    today_full = today.strftime('%Y-%m-%d')
    
    ticker_items = []
    
    # 1. Get Today's Poojs from Transactions
    pooja_df = get_data("""
        SELECT 
            CASE WHEN t.family_id = 0 THEN t.guest_name ELSE f.head_name END as name,
            s.service_name 
        FROM transactions t
        LEFT JOIN families f ON t.family_id = f.id
        JOIN services s ON t.service_id = s.id
        WHERE date(t.date) = ?
    """, (today_full,))
    for _, row in pooja_df.iterrows():
        ticker_items.append(f"🙏 Today's Pooja: {row['service_name']} by {row['name']}")
    
    # 2. Get Special Yearly Pooja Reminders (from Head & Members)
    f_p_reminders = get_data("SELECT head_name, yearly_pooja_date FROM families")
    for _, row in f_p_reminders.iterrows():
        if row['yearly_pooja_date'] and str(row['yearly_pooja_date'])[5:10] == today_md:
            ticker_items.append(f"🙏 Yearly Special Pooja Reminder for Head: {row['head_name']}!")

    p_reminders = get_data("SELECT member_name, yearly_pooja_date FROM members")
    for _, row in p_reminders.iterrows():
        if row['yearly_pooja_date'] and str(row['yearly_pooja_date'])[5:10] == today_md:
            ticker_items.append(f"🙏 Yearly Special Pooja Reminder for: {row['member_name']}!")

    # 3. Get Birthdays
    f_bday = get_data("SELECT head_name, dob FROM families")
    for _, row in f_bday.iterrows():
        if row['dob'] and str(row['dob'])[5:10] == today_md:
            ticker_items.append(f"🎂 Happy Birthday to Devotee Head: {row['head_name']}!")
            
    m_bday = get_data("SELECT member_name, dob FROM members")
    for _, row in m_bday.iterrows():
        if row['dob'] and str(row['dob'])[5:10] == today_md:
            ticker_items.append(f"🎂 Happy Birthday to {row['member_name']}!")

    # 4. Get Wedding Anniversaries
    f_anniv = get_data("SELECT head_name, wedding_date FROM families")
    for _, row in f_anniv.iterrows():
        if row['wedding_date'] and str(row['wedding_date'])[5:10] == today_md:
            ticker_items.append(f"🎉 Happy Wedding Anniversary to {row['head_name']} & Family!")
            
    m_anniv = get_data("SELECT member_name, wedding_date FROM members")
    for _, row in m_anniv.iterrows():
        if row['wedding_date'] and str(row['wedding_date'])[5:10] == today_md:
            ticker_items.append(f"🎉 Happy Wedding Anniversary to {row['member_name']}!")

    if not ticker_items:
        scrolling_text = "✨ Welcome to Sree Bhadreshwari Amman Temple Management System. May the blessings of Amman be with you. ✨"
    else:
        scrolling_text = " | ".join(ticker_items)
        scrolling_text = f"✨ {scrolling_text} ✨"

    st.markdown(f"""
        <style>
        .ticker-wrap {{
            width: 100%;
            overflow: hidden;
            background-color: #800000; 
            padding: 12px 0;
            border-bottom: 3px solid #FFD700;
            border-top: 3px solid #FFD700;
            margin-bottom: 25px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .ticker {{
            display: inline-block;
            white-space: nowrap;
            animation: marquee 45s linear infinite;
            color: #FFD700;
            font-weight: bold;
            font-size: 18px;
            padding-left: 100%;
            will-change: transform;
        }}
        @keyframes marquee {{
            0% {{ transform: translateX(0); }}
            100% {{ transform: translateX(-200%); }}
        }}
        </style>
        <div class="ticker-wrap">
            <div class="ticker">{scrolling_text} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; {scrolling_text}</div>
        </div>
    """, unsafe_allow_html=True)

# --- APP INIT ---
init_db()
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "Home Dashboard"
if 'new_family_id' not in st.session_state: st.session_state.new_family_id = None

if not st.session_state.logged_in:
    login_page()
    render_footer() 
    st.stop()

# GLOBAL GOLDEN METAL BACKGROUND CSS
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #bf953f 0%, #fcf6ba 25%, #b38728 50%, #fbf5b7 75%, #aa771c 100%);
        background-attachment: fixed;
    }
    .block-container {
        background-color: rgba(255, 255, 255, 0.05);
        border-radius: 0px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- MODULES ---

if st.session_state.current_page == "Home Dashboard":
    page_header()
    render_navigation_bar()
    render_news_ticker()
    st.title(f"Welcome, {st.session_state.username.title()}")
    
    # Financial Date Strings for SQLite filtering
    today_str = date.today().strftime('%Y-%m-%d')
    week_str = date.today().strftime('%Y-%W')
    month_str = date.today().strftime('%Y-%m')
    year_str = date.today().strftime('%Y')

    inc_today = get_data("SELECT SUM(amount) FROM transactions WHERE date(date) = ?", (today_str,)).iloc[0,0] or 0
    inc_week = get_data("SELECT SUM(amount) FROM transactions WHERE strftime('%Y-%W', date) = ?", (week_str,)).iloc[0,0] or 0
    inc_month = get_data("SELECT SUM(amount) FROM transactions WHERE strftime('%Y-%m', date) = ?", (month_str,)).iloc[0,0] or 0
    inc_year = get_data("SELECT SUM(amount) FROM transactions WHERE strftime('%Y', date) = ?", (year_str,)).iloc[0,0] or 0

    exp_today = get_data("SELECT SUM(amount) FROM users_expenses WHERE payment_date = ?", (today_str,)).iloc[0,0] or 0
    exp_week = get_data("SELECT SUM(amount) FROM users_expenses WHERE strftime('%Y-%W', payment_date) = ?", (week_str,)).iloc[0,0] or 0
    exp_month = get_data("SELECT SUM(amount) FROM users_expenses WHERE strftime('%Y-%m', payment_date) = ?", (month_str,)).iloc[0,0] or 0
    exp_year = get_data("SELECT SUM(amount) FROM users_expenses WHERE strftime('%Y', payment_date) = ?", (year_str,)).iloc[0,0] or 0
    
    df_fam = get_data("SELECT id FROM families")
    
    st.markdown("<style>[data-testid='stMetricValue'] { color: #800000; font-weight: bold; }</style>", unsafe_allow_html=True)
    
    st.subheader("💰 Income Summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Today's Income", f"₹ {inc_today:,.2f}")
    c2.metric("Weekly Income", f"₹ {inc_week:,.2f}")
    c3.metric("Monthly Income", f"₹ {inc_month:,.2f}")
    c4.metric("Yearly Income", f"₹ {inc_year:,.2f}")

    st.subheader("💸 Expense Summary")
    e1, e2, e3, e4 = st.columns(4)
    e1.metric("Today's Expenses", f"₹ {exp_today:,.2f}")
    e2.metric("Weekly Expenses", f"₹ {exp_week:,.2f}")
    e3.metric("Monthly Expenses", f"₹ {exp_month:,.2f}")
    e4.metric("Yearly Expenses", f"₹ {exp_year:,.2f}")

    st.divider()
    
    # MOVED EVENT CALENDAR TO HOME DASHBOARD
    st.subheader("📅 Daily Event Calendar")
    today_md = date.today().strftime('%m-%d')
    
    col_cal1, col_cal2 = st.columns(2)
    
    with col_cal1:
        st.markdown("#### 🔔 Today's Special Reminders")
        # Head Reminders
        f_reminders_df = get_data("SELECT head_name, phone, yearly_pooja_date FROM families")
        today_f_poojas = f_reminders_df[f_reminders_df['yearly_pooja_date'].str.contains(today_md, na=False)]
        
        # Member Reminders - Join with families to get primary phone
        reminders_df = get_data("""
            SELECT m.member_name, m.yearly_pooja_date, f.phone 
            FROM members m 
            JOIN families f ON m.family_id = f.id
        """)
        today_poojas = reminders_df[reminders_df['yearly_pooja_date'].str.contains(today_md, na=False)]
        
        if not today_f_poojas.empty or not today_poojas.empty:
            for _, row in today_f_poojas.iterrows():
                st.warning(f"🙏 **Yearly Pooja (Head):** {row['head_name']} ({row['phone']})")
            for _, row in today_poojas.iterrows():
                st.warning(f"🙏 **Yearly Pooja:** {row['member_name']} ({row['phone']})")
        else:
            st.info("No special pooja reminders for today.")
            
    with col_cal2:
        st.markdown("#### 🎂 Celebrations Today")
        # Fetch birthdays and anniversaries for Head & Members
        head_dates = get_data("SELECT head_name, phone, dob, wedding_date FROM families")
        mem_dates = get_data("""
            SELECT m.member_name, m.dob, m.wedding_date, f.phone 
            FROM members m 
            JOIN families f ON m.family_id = f.id
        """)
        
        found_celeb = False
        
        # Checking Heads
        for _, row in head_dates.iterrows():
            if row['dob'] and str(row['dob'])[5:10] == today_md:
                st.success(f"🎂 **Birthday (Head):** {row['head_name']} ({row['phone']})")
                found_celeb = True
            if row['wedding_date'] and str(row['wedding_date'])[5:10] == today_md:
                st.success(f"🎉 **Wedding Anniversary (Head):** {row['head_name']} ({row['phone']})")
                found_celeb = True
        
        # Checking Members
        for _, row in mem_dates.iterrows():
            if row['dob'] and str(row['dob'])[5:10] == today_md:
                st.success(f"🎂 **Birthday:** {row['member_name']} ({row['phone']})")
                found_celeb = True
            if row['wedding_date'] and str(row['wedding_date'])[5:10] == today_md:
                st.success(f"🎉 **Wedding Anniversary:** {row['member_name']} ({row['phone']})")
                found_celeb = True
                
        if not found_celeb:
            st.info("No birthdays or anniversaries today.")

    st.divider()
    st.metric("Total Enrolled Devotees", len(df_fam))

elif st.session_state.current_page == "Enroll":
    page_header()
    render_navigation_bar()
    st.header("Devotee Enrollment")
    
    tab_single, tab_bulk = st.tabs(["📝 Manual Entry", "📥 Bulk Upload (Excel)"])
    
    with tab_single:
        with st.expander("Step 1: Family Head Details", expanded=(st.session_state.new_family_id is None)):
            if st.session_state.new_family_id is None:
                with st.form("add_family_head_form"):
                    col_info, col_photo = st.columns([2, 1])
                    with col_info:
                        head_name = st.text_input("Family Head Name *")
                        phone = st.text_input("Primary Phone Number *")
                        whatsapp = st.text_input("WhatsApp Number")
                        address = st.text_area("Residential Address")
                    with col_photo:
                        st.write("📷 **Upload Photo**")
                        photo_file = st.file_uploader("Select Profile Image", type=['jpg', 'jpeg', 'png'])
                        if photo_file: st.image(photo_file, width=150, caption="Preview")
                    st.divider()
                    col_d, col_w, col_n, col_p = st.columns(4)
                    with col_d: head_dob = st.date_input("Date of Birth", value=None, min_value=MIN_DATE, max_value=MAX_DATE)
                    with col_w: head_wedding = st.date_input("Wedding Anniversary", value=None, min_value=MIN_DATE, max_value=MAX_DATE)
                    with col_n: head_natchathiram = st.selectbox("Select Star", [""] + NATCHATHIRAM_OPTIONS)
                    with col_p: head_pooja = st.date_input("Yearly Pooja Date", value=None)
                    
                    if st.form_submit_button("Save Family Head"):
                        if head_name and phone:
                            photo_str = image_to_base64(photo_file)
                            save_family_head(head_name, head_dob, head_wedding, head_natchathiram, address, phone, whatsapp, photo_str, head_pooja)
                        else: st.error("Name and Phone are mandatory.")
            else:
                st.info(f"Adding members for: **{st.session_state.new_family_head}**")
                if st.button("New Enrollment Process"):
                    st.session_state.new_family_id = None
                    st.rerun()
        
        if st.session_state.new_family_id:
            st.subheader("Step 2: Add Family Members")
            with st.form("add_member_form", clear_on_submit=True):
                col_nm1, col_nm2 = st.columns(2)
                with col_nm1: m_name = st.text_input("Member Name")
                with col_nm2: m_rel = st.selectbox("Relationship", [""] + RELATIONSHIP_OPTIONS)
                
                col_cnt1, col_cnt2 = st.columns(2)
                with col_cnt1: m_phone = st.text_input("Member Mobile No.")
                with col_cnt2: m_wa = st.text_input("Member WhatsApp No.")
                
                col_md, col_mw, col_mn, col_mp = st.columns(4)
                with col_md: mdob = st.date_input("DOB", value=None, min_value=MIN_DATE, max_value=MAX_DATE)
                with col_mw: mwedding = st.date_input("Anniversary", value=None, min_value=MIN_DATE, max_value=MAX_DATE)
                with col_mn: mnatch = st.selectbox("Star", [""] + NATCHATHIRAM_OPTIONS)
                with col_mp: mpooja = st.date_input("Yearly Pooja Reminder", value=None)
                
                if st.form_submit_button("Add Member"):
                    if m_name:
                        run_query("""INSERT INTO members 
                                     (family_id, member_name, relationship, dob, wedding_date, natchathiram, yearly_pooja_date, phone, whatsapp) 
                                     VALUES (?,?,?,?,?,?,?,?,?)""",
                                  (st.session_state.new_family_id, m_name, m_rel, mdob, mwedding, mnatch, mpooja, m_phone, m_wa))
                        st.success(f"Member {m_name} added!")
                    else: st.error("Name required.")

    with tab_bulk:
        st.subheader("Import Devotees from Excel")
        st.info("""
            **Required Column Headers in Excel:**
            `head_name`, `dob`, `wedding_date`, `natchathiram`, `address`, `phone`, `whatsapp`, `yearly_pooja_date`
        """)
        uploaded_file = st.file_uploader("Upload .xlsx or .xls file", type=["xlsx", "xls"])
        
        if uploaded_file:
            try:
                bulk_df = pd.read_excel(uploaded_file).fillna("")
                st.dataframe(bulk_df.head(), use_container_width=True)
                if st.button("🚀 Confirm & Import All Records"):
                    import_count = 0
                    for _, row in bulk_df.iterrows():
                        if str(row.get('head_name', '')).strip() != "":
                            run_query("""
                                INSERT INTO families (head_name, dob, wedding_date, natchathiram, address, phone, whatsapp, photo, yearly_pooja_date) 
                                VALUES (?,?,?,?,?,?,?,?,?)
                            """, (
                                str(row.get('head_name', '')),
                                str(row.get('dob', '')),
                                str(row.get('wedding_date', '')),
                                str(row.get('natchathiram', '')),
                                str(row.get('address', '')),
                                str(row.get('phone', '')),
                                str(row.get('whatsapp', '')),
                                "",
                                str(row.get('yearly_pooja_date', ''))
                            ))
                            import_count += 1
                    st.success(f"Successfully imported {import_count} records!")
            except Exception as e:
                st.error(f"Error processing file: {e}")

elif st.session_state.current_page == "Search":
    page_header()
    render_navigation_bar()
    st.header("Search & Manage Devotees")
    search_term = st.text_input("Search by Name or Mobile No.")
    query = "SELECT * FROM families"
    if search_term: query += f" WHERE head_name LIKE '%{search_term}%' OR phone LIKE '%{search_term}%'"
    df = get_data(query)
    if not df.empty:
        for idx, row in df.iterrows():
            with st.container():
                c_img, c_detail = st.columns([1, 6])
                with c_img:
                    img_data = base64_to_image(row['photo'])
                    if img_data: st.image(img_data, width=100)
                    else: st.markdown("<div style='width:100px; height:100px; background:#f0f0f0; border:1px solid #ddd; display:flex; align-items:center; justify-content:center; border-radius:0px;'>👤</div>", unsafe_allow_html=True)
                with c_detail:
                    st.subheader(row['head_name'])
                    st.write(f"📞 {row['phone']} | ⭐ {row['natchathiram'] or 'N/A'} | 🏠 {row['address']}")
                    
                    # --- EDIT FAMILY HEAD ---
                    with st.expander(f"⚙️ Edit Family Head: {row['head_name']}"):
                        with st.form(f"edit_form_{row['id']}"):
                            ce1, ce2 = st.columns(2)
                            with ce1:
                                new_name = st.text_input("Name", value=row['head_name'])
                                new_phone = st.text_input("Phone", value=row['phone'])
                                new_wa = st.text_input("WhatsApp", value=row['whatsapp'])
                            with ce2:
                                current_dob = datetime.strptime(row['dob'], '%Y-%m-%d').date() if row['dob'] and str(row['dob']) not in ['None', ''] else None
                                current_wed = datetime.strptime(row['wedding_date'], '%Y-%m-%d').date() if row['wedding_date'] and str(row['wedding_date']) not in ['None', ''] else None
                                current_pooja = datetime.strptime(row['yearly_pooja_date'], '%Y-%m-%d').date() if row['yearly_pooja_date'] and str(row['yearly_pooja_date']) not in ['None', ''] else None
                                
                                new_dob = st.date_input("DOB", value=current_dob, min_value=MIN_DATE, max_value=MAX_DATE)
                                new_wed = st.date_input("Anniversary", value=current_wed, min_value=MIN_DATE, max_value=MAX_DATE)
                                new_pooja = st.date_input("Yearly Pooja Date", value=current_pooja)
                                new_star = st.selectbox("Star", NATCHATHIRAM_OPTIONS, index=NATCHATHIRAM_OPTIONS.index(row['natchathiram']) if row['natchathiram'] in NATCHATHIRAM_OPTIONS else 0)
                            
                            new_addr = st.text_area("Address", value=row['address'])
                            new_photo_file = st.file_uploader("Update Photo", type=['jpg','png'], key=f"photo_{row['id']}")
                            
                            if st.form_submit_button("Save Changes to Head"):
                                update_photo = image_to_base64(new_photo_file) if new_photo_file else row['photo']
                                run_query("""
                                    UPDATE families SET 
                                    head_name=?, dob=?, wedding_date=?, natchathiram=?, address=?, phone=?, whatsapp=?, photo=?, yearly_pooja_date=? 
                                    WHERE id=?
                                """, (new_name, new_dob, new_wed, new_star, new_addr, new_phone, new_wa, update_photo, new_pooja, row['id']))
                                st.success("Head Profile Updated!")
                                st.rerun()
                        
                        # DELETE FAMILY HEAD OPTION
                        st.markdown("---")
                        st.write("🗑️ **Danger Zone**")
                        if st.button(f"Delete Devotee Account: {row['head_name']}", key=f"del_head_{row['id']}"):
                            # Delete associated members first
                            run_query("DELETE FROM members WHERE family_id=?", (row['id'],))
                            # Delete transactions or link them to Guest (Better to just delete for clean management)
                            run_query("DELETE FROM families WHERE id=?", (row['id'],))
                            st.warning(f"Account for {row['head_name']} and all associated family members deleted.")
                            st.rerun()

                    # --- EDIT FAMILY MEMBERS ---
                    with st.expander(f"👨‍👩‍👧 Manage Family Members ({row['head_name']})"):
                        members_df = get_data("SELECT * FROM members WHERE family_id = ?", (row['id'],))
                        if members_df.empty:
                            st.info("No members recorded for this family.")
                        else:
                            for midx, mrow in members_df.iterrows():
                                st.markdown(f"**Member:** {mrow['member_name']} ({mrow['relationship']})")
                                with st.form(f"edit_member_{mrow['id']}"):
                                    mce1, mce2 = st.columns(2)
                                    with mce1:
                                        m_new_name = st.text_input("Member Name", value=mrow['member_name'])
                                        m_new_rel = st.selectbox("Relationship", RELATIONSHIP_OPTIONS, index=RELATIONSHIP_OPTIONS.index(mrow['relationship']) if mrow['relationship'] in RELATIONSHIP_OPTIONS else 0)
                                        m_new_star = st.selectbox("Star", NATCHATHIRAM_OPTIONS, index=NATCHATHIRAM_OPTIONS.index(mrow['natchathiram']) if mrow['natchathiram'] in NATCHATHIRAM_OPTIONS else 0, key=f"star_m_{mrow['id']}")
                                        m_new_phone = st.text_input("Mobile No.", value=mrow['phone'] if mrow['phone'] else "")
                                        m_new_wa = st.text_input("WhatsApp No.", value=mrow['whatsapp'] if mrow['whatsapp'] else "")
                                    with mce2:
                                        m_curr_dob = datetime.strptime(mrow['dob'], '%Y-%m-%d').date() if mrow['dob'] and str(mrow['dob']) not in ['None', ''] else None
                                        m_curr_wed = datetime.strptime(mrow['wedding_date'], '%Y-%m-%d').date() if mrow['wedding_date'] and str(mrow['wedding_date']) not in ['None', ''] else None
                                        m_curr_pj = datetime.strptime(mrow['yearly_pooja_date'], '%Y-%m-%d').date() if mrow['yearly_pooja_date'] and str(mrow['yearly_pooja_date']) not in ['None', ''] else None
                                        
                                        m_new_dob = st.date_input("DOB", value=m_curr_dob, key=f"dob_m_{mrow['id']}")
                                        m_new_wed = st.date_input("Anniversary", value=m_curr_wed, key=f"wed_m_{mrow['id']}")
                                        m_new_pj = st.date_input("Yearly Pooja", value=m_curr_pj, key=f"pj_m_{mrow['id']}")
                                    
                                    col_eb1, col_eb2 = st.columns([1,1])
                                    if col_eb1.form_submit_button("Update Member"):
                                        run_query("""
                                            UPDATE members SET member_name=?, relationship=?, dob=?, wedding_date=?, natchathiram=?, yearly_pooja_date=?, phone=?, whatsapp=?
                                            WHERE id=?
                                        """, (m_new_name, m_new_rel, m_new_dob, m_new_wed, m_new_star, m_new_pj, m_new_phone, m_new_wa, mrow['id']))
                                        st.success(f"Member {m_new_name} Updated!")
                                        st.rerun()
                                    if col_eb2.form_submit_button("🗑️ Remove Member"):
                                        run_query("DELETE FROM members WHERE id=?", (mrow['id'],))
                                        st.warning("Member removed.")
                                        st.rerun()
                                st.divider()

                st.divider()

elif st.session_state.current_page == "Billing":
    page_header()
    render_navigation_bar()
    st.header("Billing Desk")
    
    billing_mode = st.radio("Billing Mode", ["Enrolled Devotee", "Guest Devotee"], horizontal=True)
    
    col1, col2 = st.columns([2, 1])
    with col1:
        selected_name = ""
        selected_address = ""
        selected_id = 0
        selected_wa = "" # Reset whatsapp number

        if billing_mode == "Enrolled Devotee":
            fams = get_data("SELECT id, head_name, phone, address, whatsapp FROM families")
            if not fams.empty:
                fam_list = {f"{r['head_name']} ({r['phone']})": r for _, r in fams.iterrows()}
                sel_fam_name = st.selectbox("Choose Devotee", list(fam_list.keys()))
                devotee = fam_list[sel_fam_name]
                selected_name = devotee['head_name']
                selected_address = devotee['address']
                selected_id = devotee['id']
                # Favor whatsapp field, fallback to phone
                selected_wa = str(devotee['whatsapp']).strip() if str(devotee['whatsapp']).strip() else str(devotee['phone']).strip()
            else:
                st.warning("Enroll a devotee first or use Guest mode.")
        else:
            selected_name = st.text_input("Guest Name *")
            selected_address = st.text_area("Guest Address")
            selected_wa = st.text_input("Guest WhatsApp No.")
            selected_id = 0 # Guest indicator

        services = get_data("SELECT id, service_name, price FROM services")
        if not services.empty:
            serv_dict = {r['service_name']: r for _, r in services.iterrows()}
            sel_serv = st.selectbox("Select Pooja/Service", list(serv_dict.keys()))
            service = serv_dict[sel_serv]
            
            st.divider()
            c_m1, c_m2 = st.columns(2)
            with c_m1: manual_no = st.text_input("Manual Bill No.")
            with c_m2: book_no = st.text_input("Bill Book No.")
            st.metric("Total Amount Payable", f"₹ {service['price']:,.2f}")
            
            if st.button("Generate Receipt"):
                if billing_mode == "Guest Devotee" and not selected_name:
                    st.error("Please enter Guest Name.")
                elif billing_mode == "Enrolled Devotee" and selected_id == 0:
                    st.error("Please select a devotee.")
                else:
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    run_query("""
                        INSERT INTO transactions (family_id, service_id, amount, date, manual_bill_no, bill_book_no, guest_name, guest_address, guest_whatsapp) 
                        VALUES (?,?,?,?,?,?,?,?,?)
                    """, (selected_id, service['id'], service['price'], now_str, manual_no, book_no, 
                          selected_name if selected_id == 0 else "", 
                          selected_address if selected_id == 0 else "",
                          selected_wa))
                    
                    last_id = get_data("SELECT max(id) as id FROM transactions").iloc[0]['id']
                    pdf = generate_pdf(last_id, selected_name, selected_address, sel_serv, service['price'], now_str, manual_no, book_no)
                    
                    st.success("Receipt generated successfully!")
                    
                    # Action Buttons Row
                    col_act1, col_act2 = st.columns(2)
                    with col_act1:
                        st.download_button("📥 Download PDF Receipt", pdf, f"Receipt_{last_id}.pdf", "application/pdf", use_container_width=True)
                    
                    with col_act2:
                        if selected_wa and str(selected_wa).strip():
                            # Construct and Encode WhatsApp Message
                            whatsapp_msg = (
                                f"🙏 *{TEMPLE_NAME_FULL}*\n\n"
                                f"Namaste *{selected_name}*,\n"
                                f"Receipt No: *#{last_id}*\n"
                                f"Manual Bill No: *{manual_no if manual_no else 'N/A'}*\n"
                                f"Bill Book No: *{book_no if book_no else 'N/A'}*\n"
                                f"Service: *{sel_serv}*\n"
                                f"Amount Paid: *₹{service['price']:,.2f}*\n"
                                f"Date: *{now_str}*\n\n"
                                f"May the blessings of Amman be with you. ✨"
                            )
                            encoded_msg = whatsapp_msg.replace(' ', '%20').replace('\n', '%0A').replace('#', '%23').replace('*', '%2A')
                            
                            # Clean phone number (keep digits only)
                            clean_wa = "".join(filter(str.isdigit, str(selected_wa)))
                            if len(clean_wa) == 10: clean_wa = "91" + clean_wa
                            
                            wa_link = f"https://wa.me/{clean_wa}?text={encoded_msg}"
                            st.link_button("📲 Send to WhatsApp", wa_link, use_container_width=True)
                        else:
                            st.info("Enter a WhatsApp No. above to enable mobile sharing.")
        else:
            st.error("No services found. Please add services in Settings first.")
    
    with col2:
        st.subheader("Recent History")
        history_df = get_data("""
            SELECT t.id as ID, t.amount as Amt,
            CASE WHEN t.family_id = 0 THEN t.guest_name ELSE f.head_name END as Devotee
            FROM transactions t
            LEFT JOIN families f ON t.family_id = f.id
            ORDER BY t.id DESC LIMIT 10
        """)
        st.dataframe(history_df, hide_index=True, use_container_width=True)

elif st.session_state.current_page == "Expenses":
    page_header()
    render_navigation_bar()
    st.header("Expense Management")
    tab_add, tab_view = st.tabs(["Record Expense", "Expense History"])
    with tab_add:
        with st.form("add_expense_form", clear_on_submit=True):
            exp_name = st.text_input("Expense Title *")
            exp_type = st.selectbox("Category", EXPENSE_TYPES)
            exp_amt = st.number_input("Amount (₹) *", min_value=0.0)
            exp_date = st.date_input("Date", value=date.today())
            exp_method = st.selectbox("Payment Method", PAYMENT_METHODS)
            exp_status = st.selectbox("Status", PAYMENT_STATUS)
            exp_vouch = st.text_input("Voucher / Bill No.")
            exp_desc = st.text_area("Additional Description")
            if st.form_submit_button("Save Expense"):
                if exp_name and exp_amt > 0:
                    run_query("INSERT INTO users_expenses (expense_name, description, expense_type, amount, payment_method, status, payment_date, voucher_no) VALUES (?,?,?,?,?,?,?,?)",
                              (exp_name, exp_desc, exp_type, exp_amt, exp_method, exp_status, exp_date, exp_vouch))
                    st.success("Expense recorded successfully!")
                else: st.error("Title and amount required.")
    with tab_view:
        df_exp = get_data("SELECT expense_name, expense_type, amount, payment_date, status FROM users_expenses ORDER BY payment_date DESC")
        st.dataframe(df_exp, use_container_width=True)

elif st.session_state.current_page == "Reports":
    page_header()
    render_navigation_bar()
    st.header("Financial & Administrative Reports")
    report_type = st.radio("Select Period:", ["Daily", "Monthly", "Custom Date Range"], horizontal=True)
    start_date, end_date = None, None
    if report_type == "Daily":
        start_date = st.date_input("Start Date", value=date.today()); end_date = start_date
    elif report_type == "Monthly":
        c_y, c_m = st.columns(2)
        with c_y: year = st.number_input("Year", value=date.today().year, min_value=2000)
        with c_m: sel_month = st.selectbox("Month", range(1,13), index=date.today().month - 1)
        start_date = date(year, sel_month, 1); last_day = calendar.monthrange(year, sel_month)[1]; end_date = date(year, sel_month, last_day)
    else:
        c_s, c_e = st.columns(2)
        with c_s: start_date = st.date_input("From Date", value=date.today())
        with c_e: end_date = st.date_input("To Date", value=date.today())
    income_query = f"SELECT SUM(amount) FROM transactions WHERE date BETWEEN '{start_date} 00:00:00' AND '{end_date} 23:59:59'"
    total_income = get_data(income_query).iloc[0,0] or 0
    st.metric("Aggregate Income", f"₹ {total_income:,.2f}")
    if total_income > 0:
        df_income_detail = get_data(f"""
            SELECT 
                t.date AS Date, 
                CASE WHEN t.family_id = 0 THEN t.guest_name ELSE f.head_name END AS Devotee, 
                s.service_name AS Service, 
                t.amount AS Amount, 
                t.manual_bill_no AS Physical_No, 
                t.bill_book_no AS Book_No
            FROM transactions t 
            JOIN services s ON t.service_id = s.id 
            LEFT JOIN families f ON t.family_id = f.id
            WHERE t.date BETWEEN '{start_date} 00:00:00' AND '{end_date} 23:59:59' 
            ORDER BY t.date DESC
        """)
        st.dataframe(df_income_detail, use_container_width=True)
        col_d1, col_d2, col_d3 = st.columns(3)
        col_d1.download_button("Download CSV", df_income_detail.to_csv(index=False).encode('utf-8'), f"Income_{start_date}.csv", "text/csv", use_container_width=True)
        
        excel_data = to_excel(df_income_detail)
        if excel_data:
            col_d2.download_button("Download Excel", excel_data, f"Income_{start_date}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
        pdf_data = generate_income_pdf(df_income_detail[['Date', 'Devotee', 'Service', 'Amount']], f"Income: {start_date} to {end_date}", total_income)
        col_d3.download_button("Download PDF", pdf_data, f"Income_{start_date}.pdf", "application/pdf", use_container_width=True)

elif st.session_state.current_page == "Assets":
    page_header()
    render_navigation_bar()
    st.header("Asset Management")
    with st.expander("Add New Temple Asset"):
        with st.form("add_asset_form", clear_on_submit=True):
            a_name = st.text_input("Asset Name *")
            a_desc = st.text_input("Description")
            a_val = st.number_input("Est. Value (₹)", min_value=0.0)
            a_qty = st.number_input("Quantity", min_value=1)
            if st.form_submit_button("Register Asset"):
                if a_name:
                    run_query("INSERT INTO assets (asset_name, description, value, quantity) VALUES (?,?,?,?)", (a_name, a_desc, a_val, a_qty))
                    st.success("Asset registered.")
                else: st.error("Asset name required.")
    df_assets = get_data("SELECT id, asset_name, description, value, quantity FROM assets")
    st.dataframe(df_assets, use_container_width=True)

elif st.session_state.current_page == "Settings":
    page_header()
    render_navigation_bar()
    st.header("System Settings")
    tab_add_s, tab_manage_s = st.tabs(["Add Service", "Manage Services"])
    with tab_add_s:
        with st.form("add_service_form", clear_on_submit=True):
            s_name = st.text_input("Service Name *")
            s_price = st.number_input("Price (₹) *", min_value=0.0)
            if st.form_submit_button("Save New Service"):
                if s_name:
                    run_query("INSERT INTO services (service_name, price) VALUES (?,?)", (s_name, s_price))
                    st.success("Service added.")
                else: st.error("Name required.")
    with tab_manage_s:
        df_serv = get_data("SELECT id, service_name, price FROM services")
        st.dataframe(df_serv, use_container_width=True)
        if not df_serv.empty:
            sel_id = st.selectbox("Select ID to Delete", df_serv['id'].tolist())
            if st.button("🚨 Delete Selected Service"):
                run_query("DELETE FROM services WHERE id=?", (sel_id,))
                st.warning("Deleted.")
                st.rerun()

elif st.session_state.current_page == "Users":
    page_header()
    render_navigation_bar()
    st.header("User Management")
    if st.session_state.role != ADMIN_ROLE: st.error("Access Restricted.")
    else:
        with st.form("create_user_form", clear_on_submit=True):
            new_u = st.text_input("New Username *")
            new_p = st.text_input("New Password *", type="password")
            new_r = st.selectbox("Assign Role", [USER_ROLE, ADMIN_ROLE])
            if st.form_submit_button("Create Account"):
                if new_u and new_p:
                    try:
                        run_query("INSERT INTO users (username, password_hash, role) VALUES (?,?,?)", (new_u, hash_password(new_p), new_r))
                        st.success(f"User '{new_u}' created!")
                    except: st.error("Username exists.")
                else: st.error("Required fields missing.")
        st.divider()
        st.subheader("System Accounts")
        st.dataframe(get_data("SELECT id, username, role FROM users"), use_container_width=True)

render_footer()