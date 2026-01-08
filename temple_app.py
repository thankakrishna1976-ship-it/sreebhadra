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

# Default/Fallback Options
DEFAULT_EXPENSE_TYPES = ['Pooja Items', 'Maintenance/Repairs', 'Salary/Dakshina', 'Electricity/Water', 'Annadanam/Food', 'Construction', 'Festivals', 'Administrative', 'Other']
RELATIONSHIP_OPTIONS = ['Wife', 'Son', 'Daughter', 'Mother', 'Father', 'Grand Father', 'Grand Mother', 'Guardian', 'Other']
NATCHATHIRAM_OPTIONS = ['Ashwini', 'Bharani', 'Karthigai', 'Rohini', 'Mrigasiram', 'Thiruvathirai', 'Punarpoosam', 'Poosam', 'Ayilyam', 'Magam', 'Poorvam', 'Uthiram', 'Hastham', 'Chithirai', 'Swathi', 'Visakam', 'Anusham', 'Kettai', 'Moolam', 'Pooradam', 'Uthiradam', 'Thiruvonam', 'Avittam', 'Sathayam', 'Poorattathi', 'Uthirattathi', 'Revathi']
PAYMENT_METHODS = ['Cash', 'UPI / GPay', 'Bank Transfer', 'Cheque', 'Card']
PAYMENT_STATUS = ['Paid', 'Pending', 'Partial']

MIN_DATE = date(1940, 1, 1)
MAX_DATE = date(2040, 12, 31)

# Menu Keys for Rights Management
ALL_MENU_KEYS = ["Home Dashboard", "Enroll", "Search", "Billing", "Expenses", "Reports", "Assets", "Settings"]

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

def get_supabase_df(table_name, select="*"):
    try:
        response = supabase.table(table_name).select(select).execute()
        return pd.DataFrame(response.data)
    except Exception as e:
        st.error(f"Database Fetch Error: {e}")
        return pd.DataFrame()

def get_data(table_name, select="*"):
    return get_supabase_df(table_name, select)

# --- UTILITY FUNCTIONS ---

def format_date_for_db(val):
    """Handles Pandas NaT and empty strings to ensure valid SQL DATE or NULL."""
    if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nat":
        return None
    try:
        d_obj = pd.to_datetime(val)
        return d_obj.strftime('%Y-%m-%d')
    except:
        return None

def format_date_for_ui(val):
    """Converts YYYY-MM-DD from database to DD/MM/YYYY for tabular display."""
    if not val or str(val).lower() in ["none", "nat", ""]:
        return ""
    try:
        d_obj = pd.to_datetime(val)
        return d_obj.strftime('%d/%m/%Y')
    except:
        return str(val)

def safe_date_convert(val):
    """Safely converts string or object to date object for st.date_input."""
    if not val or str(val).lower() in ["none", "nat", ""]:
        return None
    try:
        return pd.to_datetime(val).date()
    except:
        return None

def get_base64_of_bin_file(bin_file):
    try:
        if os.path.exists(bin_file):
            with open(bin_file, 'rb') as f:
                data = f.read()
            return base64.b64encode(data).decode()
    except: pass
    return None

def image_to_base64(image_file):
    if image_file is not None:
        return base64.b64encode(image_file.getvalue()).decode()
    return ""

def base64_to_image(base64_str):
    if base64_str:
        try: return io.BytesIO(base64.b64decode(base64_str))
        except: return None
    return None

def to_excel(df):
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Report')
        return output.getvalue()
    except Exception:
        return None

def generate_financial_pdf(income_df, expense_df, title, t_inc, t_exp, t_net):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=50, bottomMargin=30)
    styles = getSampleStyleSheet()
    story = [Paragraph(TEMPLE_NAME_FULL, styles['Title']), Paragraph(title, styles['h3']), Spacer(1, 12)]
    
    # Summary Table
    summary_data = [
        ["Total Income", f"₹ {t_inc:,.2f}"],
        ["Total Expenses", f"₹ {t_exp:,.2f}"],
        ["Net Profit", f"₹ {t_net:,.2f}"]
    ]
    st_table = Table(summary_data, colWidths=[150, 150])
    st_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold')
    ]))
    story.append(st_table)
    story.append(Spacer(1, 24))
    
    # Income Section
    story.append(Paragraph("Income Details", styles['h4']))
    if not income_df.empty:
        inc_data = [income_df.columns.tolist()] + income_df.values.tolist()
        t_inc = Table(inc_data, hAlign='LEFT', repeatRows=1)
        t_inc.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#800000')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.gold),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        story.append(t_inc)
    else:
        story.append(Paragraph("No income records for this period.", styles['Normal']))
        
    story.append(Spacer(1, 24))

    # Expense Section
    story.append(Paragraph("Expense Details", styles['h4']))
    if not expense_df.empty:
        exp_data = [expense_df.columns.tolist()] + expense_df.values.tolist()
        t_exp = Table(exp_data, hAlign='LEFT', repeatRows=1)
        t_exp.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#800000')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.gold),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        story.append(t_exp)
    else:
        story.append(Paragraph("No expense records for this period.", styles['Normal']))

    doc.build(story)
    buffer.seek(0)
    return buffer

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def verify_user(username, password):
    res = supabase.table('users').select('*').eq('username', username).execute()
    if res.data:
        user_data = res.data[0]
        if user_data['password_hash'] == hash_password(password):
            # Split rights string into list
            user_rights = user_data.get('rights', 'Home Dashboard').split(',')
            return True, user_data['role'], user_rights
    return False, None, None

# --- FOOTER ---
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

# --- RECEIPT PDF ---
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
    c.drawString(180, amount_y, f"Rs. {float(amount):,.2f}/-")
    c.setFillColorRGB(0, 0, 0) 
    footer_y = amount_y - 100
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(50, footer_y, "Thank you for your offering. May the blessings of Sree Bhadreshwari Amman be upon you.")
    c.setFont("Helvetica", 10)
    c.drawString(400, footer_y - 50, "Authorized Signature")
    c.save()
    buffer.seek(0)
    return buffer

# --- LOGIN PAGE ---
def login_page():
    bg_img_base64 = get_base64_of_bin_file(BACKGROUND_PATH)
    bg_style = f"""
        <style>
        .stApp {{ 
            {"background-image: url('data:image/jpg;base64," + bg_img_base64 + "');" if bg_img_base64 else ""}
            background-size: 75% !important; background-position: center !important;
            background-repeat: no-repeat !important; background-attachment: fixed !important;
            background-color: #bf953f; 
        }}
        label {{ color: #800000 !important; font-weight: 900 !important; font-size: 20px !important; text-align: left !important; }}
        .stButton>button {{ 
            width: 100%; background-color: #800000; color: #FFD700; padding: 10px;
            border-radius: 0px !important; border: 1px solid #FFD700; font-size: 16px; font-weight: bold;
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
        st.markdown("""<div style='background-color: #800000; color: #FFD700; padding: 15px; text-align: center; font-weight: bold; font-size: 24px; border: 2px solid #FFD700; margin-bottom: 20px;'>Temple Management System</div>""", unsafe_allow_html=True)
        st.markdown("""<div style='background-color: #800000; color: #FFD700; padding: 10px; text-align: center; font-weight: 900; font-size: 22px; border: 1px solid #FFD700; margin-bottom: 15px;'>Staff Login</div>""", unsafe_allow_html=True)
        username = st.text_input("Username", key="login_user")
        password = st.text_input("Password", type="password", key="login_pass")
        if st.button("SIGN IN", key="login_button"):
            success, role, rights = verify_user(username, password)
            if success:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.role = role
                st.session_state.rights = rights
                st.rerun()
            else: st.error("Invalid Username or Password.")

# --- NAVIGATION ---
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
        "Assets": {"label": "ASSETS"}, "Settings": {"label": "SETTINGS"},
    }
    
    # Filter pages based on rights (Admins get all)
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
            height: 4em; width: 100%; margin-top: -10px; transition: all 0.3s ease;
            text-transform: uppercase; box-shadow: 2px 2px 5px rgba(0,0,0,0.2);
        }
        div[data-testid="column"] .stButton>button:hover {
            background-color: #A00000; color: #FFFFFF; border-color: #FFFFFF;
            box-shadow: 0 6px 15px rgba(0,0,0,0.3); transform: translateY(-2px);
        }
        div[data-testid="column"] .stButton>button:active { background-color: #FFD700; color: #800000; }
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
        data = {
            "head_name": head_name, 
            "dob": format_date_for_db(dob),
            "wedding_date": format_date_for_db(wedding_date),
            "natchathiram": n_star, "address": address, "phone": phone,
            "whatsapp": whatsapp, "photo": photo_str,
            "yearly_pooja_date": format_date_for_db(pooja_date)
        }
        res = run_supabase_insert("families", data)
        if res and res.data:
            st.session_state.new_family_id = res.data[0]['id']
            st.session_state.new_family_head = head_name
            st.success(f"Family Head {head_name} saved successfully!")
            st.rerun()

# --- TICKER ---
def render_news_ticker():
    today = date.today()
    today_md = today.strftime('%m-%d')
    ticker_items = []
    df_fam = get_data("families")
    if not df_fam.empty:
        # Check dob exists and is not None before filtering
        b_df = df_fam[df_fam['dob'].astype(str).str.contains(today_md, na=False)]
        for _, r in b_df.iterrows(): ticker_items.append(f"🎂 Happy Birthday to Devotee Head: {r['head_name']}!")
        p_df = df_fam[df_fam['yearly_pooja_date'].astype(str).str.contains(today_md, na=False)]
        for _, r in p_df.iterrows(): ticker_items.append(f"🙏 Yearly Special Pooja Reminder for Head: {r['head_name']}!")
    scrolling_text = " | ".join(ticker_items) if ticker_items else "✨ Welcome to Sree Bhadreshwari Amman Temple Management System. ✨"
    st.markdown(f"""
        <style>
        .ticker-wrap {{ width: 100%; overflow: hidden; background-color: #800000; padding: 12px 0; border-bottom: 3px solid #FFD700; border-top: 3px solid #FFD700; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
        .ticker {{ display: inline-block; white-space: nowrap; animation: marquee 45s linear infinite; color: #FFD700; font-weight: bold; font-size: 18px; padding-left: 100%; will-change: transform; }}
        @keyframes marquee {{ 0% {{ transform: translateX(0); }} 100% {{ transform: translateX(-200%); }} }}
        </style>
        <div class="ticker-wrap"><div class="ticker">{scrolling_text} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; {scrolling_text}</div></div>
    """, unsafe_allow_html=True)

# --- APP INIT ---
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "Home Dashboard"
if 'new_family_id' not in st.session_state: st.session_state.new_family_id = None

if not st.session_state.logged_in:
    login_page()
    render_footer() 
    st.stop()

st.markdown("""
    <style>
    .stApp { background: linear-gradient(135deg, #bf953f 0%, #fcf6ba 25%, #b38728 50%, #fbf5b7 75%, #aa771c 100%); background-attachment: fixed; }
    .block-container { background-color: rgba(255, 255, 255, 0.05); border-radius: 0px; }
    </style>
    """, unsafe_allow_html=True)

# --- MODULES ---

if st.session_state.current_page == "Home Dashboard":
    page_header()
    render_navigation_bar()
    render_news_ticker()
    st.title(f"Welcome, {st.session_state.username.title()}")
    
    # Financial Overview Calculation
    today = date.today()
    start_of_week = today - timedelta(days=today.weekday())
    start_of_month = today.replace(day=1)
    start_of_year = today.replace(month=1, day=1)

    df_trans = get_data("transactions")
    df_exp_all = get_data("users_expenses")
    
    if not df_trans.empty:
        df_trans['date_obj'] = pd.to_datetime(df_trans['date']).dt.date
    if not df_exp_all.empty:
        df_exp_all['date_obj'] = pd.to_datetime(df_exp_all['payment_date']).dt.date

    def calc_finances(start_d, end_d):
        inc = 0
        exp = 0
        if not df_trans.empty:
            inc = df_trans[(df_trans['date_obj'] >= start_d) & (df_trans['date_obj'] <= end_d)]['amount'].sum()
        if not df_exp_all.empty:
            exp = df_exp_all[(df_exp_all['date_obj'] >= start_d) & (df_exp_all['date_obj'] <= end_d)]['amount'].sum()
        return inc, exp, (inc - exp)

    stats = {
        "Daily": calc_finances(today, today),
        "Weekly": calc_finances(start_of_week, today),
        "Monthly": calc_finances(start_of_month, today),
        "Yearly": calc_finances(start_of_year, today)
    }

    for label, (inc, exp, net) in stats.items():
        st.subheader(f"📊 {label} Overview")
        c1, c2, c3 = st.columns(3)
        c1.metric(f"{label} Income", f"₹ {inc:,.2f}")
        c2.metric(f"{label} Expense", f"₹ {exp:,.2f}")
        c3.metric(f"{label} Net Profit", f"₹ {net:,.2f}")
        st.divider()

    df_fam = get_data("families")
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
                        if photo_file: st.image(photo_file, width=150)
                    st.divider()
                    col_d, col_w, col_n, col_p = st.columns(4)
                    with col_d: head_dob = st.date_input("Date of Birth", value=None, min_value=MIN_DATE)
                    with col_w: head_wedding = st.date_input("Wedding Anniversary", value=None, min_value=MIN_DATE)
                    with col_n: head_natchathiram = st.selectbox("Select Star", [""] + NATCHATHIRAM_OPTIONS)
                    with col_p: head_pooja = st.date_input("Yearly Pooja Date", value=None)
                    if st.form_submit_button("Save Family Head"):
                        save_family_head(head_name, head_dob, head_wedding, head_natchathiram, address, phone, whatsapp, image_to_base64(photo_file), head_pooja)
            else:
                st.info(f"Adding members for: **{st.session_state.new_family_head}**")
                if st.button("New Enrollment Process"):
                    st.session_state.new_family_id = None
                    st.rerun()
        
        if st.session_state.new_family_id:
            st.subheader("Step 2: Add Family Members")
            with st.form("add_member_form", clear_on_submit=True):
                col_nm1, col_nm2 = st.columns(2)
                with col_nm1: m_name = st.text_input("Member Name *")
                with col_nm2: m_rel = st.selectbox("Relationship *", [""] + RELATIONSHIP_OPTIONS)
                
                col_cnt1, col_cnt2 = st.columns(2)
                with col_cnt1: m_phone = st.text_input("Member Mobile No.")
                with col_cnt2: m_wa = st.text_input("Member WhatsApp No.")
                
                col_dates1, col_dates2 = st.columns(2)
                with col_dates1: m_dob = st.date_input("Date of Birth", value=None, min_value=MIN_DATE)
                with col_dates2: m_wedding = st.date_input("Wedding Day", value=None, min_value=MIN_DATE)
                
                col_extra1, col_extra2 = st.columns(2)
                with col_extra1: mnatch = st.selectbox("Star", [""] + NATCHATHIRAM_OPTIONS)
                with col_extra2: mpooja = st.date_input("Yearly Pooja Reminder", value=None)
                
                if st.form_submit_button("Add Member"):
                    if m_name:
                        data = {
                            "family_id": st.session_state.new_family_id, 
                            "member_name": m_name, 
                            "relationship": m_rel, 
                            "phone": m_phone, 
                            "whatsapp": m_wa,
                            "dob": format_date_for_db(m_dob),
                            "wedding_date": format_date_for_db(m_wedding),
                            "natchathiram": mnatch,
                            "yearly_pooja_date": format_date_for_db(mpooja)
                        }
                        run_supabase_insert("members", data)
                        st.success(f"Member {m_name} added!")
                    else: st.error("Name required.")

    with tab_bulk:
        st.subheader("Bulk Import Devotees")
        # --- SAMPLE EXCEL DOWNLOAD ---
        st.markdown("### 📥 Download Sample Template")
        sample_data = {
            "Sl.No": [1, "", "", 2],
            "Name": ["Rajesh Kumar", "", "", "Suresh Nair"],
            "Address": ["123 Temple St, Kanjampuram", "", "", "456 Main Rd, Kanyakumari"],
            "Mobile No": ["9876543210", "", "", "9988776655"],
            "WhatsApp No": ["9876543210", "", "", "9988776655"],
            "Members": ["Rajesh Kumar", "Priya Rajesh", "Anand Rajesh", "Suresh Nair"],
            "Relationship": ["Family Head", "Wife", "Son", "Family Head"],
            "Date of Birth": ["1980-05-15", "1985-08-20", "2012-03-10", "1975-01-10"],
            "Star": ["Ashwini", "Swathi", "Revathi", "Bharani"],
            "Wedding Day": ["2010-06-12", "", "", "2005-02-14"],
            "Yearly Pooja": ["2024-11-10", "", "", "2024-03-15"]
        }
        sample_df = pd.DataFrame(sample_data)
        excel_sample = to_excel(sample_df)
        if excel_sample:
            st.download_button(
                label="📁 Download Sample Excel Sheet",
                data=excel_sample,
                file_name="temple_bulk_upload_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        st.divider()
        st.info("""
            **Excel Format Instructions:**
            1. New families start with a number in the **'Sl.No'** column.
            2. For family members, leave **'Sl.No'**, **'Name'**, **'Address'**, etc., blank.
            3. **'Relationship'** for the head should be 'Family Head'.
        """)
        
        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
        if uploaded_file:
            try:
                # Use engine='openpyxl' for .xlsx
                bulk_df = pd.read_excel(uploaded_file, engine='openpyxl')
                st.write("File Preview:")
                st.dataframe(bulk_df.head())
                
                if st.button("🚀 Process Bulk Upload"):
                    success_count = 0
                    current_fid = None
                    
                    # English headers
                    col_sl = "Sl.No"
                    col_name = "Name"
                    col_addr = "Address"
                    col_mob = "Mobile No"
                    col_wa = "WhatsApp No"
                    col_mem_name = "Members"
                    col_rel = "Relationship"
                    col_dob = "Date of Birth"
                    col_star = "Star"
                    col_wed = "Wedding Day"
                    col_pj = "Yearly Pooja"

                    for _, row in bulk_df.iterrows():
                        is_head = str(row.get(col_rel, "")).strip().lower() == "family head" or not pd.isna(row.get(col_sl))
                        
                        if is_head and not pd.isna(row.get(col_name)) and str(row.get(col_name)).strip() != "":
                            # Insert New Family Head
                            head_data = {
                                "head_name": str(row[col_name]),
                                "address": str(row[col_addr]),
                                "phone": str(row[col_mob]),
                                "whatsapp": str(row[col_wa]),
                                "natchathiram": str(row[col_star]),
                                "dob": format_date_for_db(row[col_dob]),
                                "wedding_date": format_date_for_db(row[col_wed]),
                                "yearly_pooja_date": format_date_for_db(row[col_pj])
                            }
                            h_res = run_supabase_insert("families", head_data)
                            if h_res and h_res.data:
                                current_fid = h_res.data[0]['id']
                                success_count += 1
                        
                        elif current_fid and not pd.isna(row.get(col_mem_name)) and str(row.get(col_mem_name)).strip() != "":
                            # Insert Family Member linked to current head
                            m_data = {
                                "family_id": current_fid,
                                "member_name": str(row[col_mem_name]),
                                "relationship": str(row[col_rel]),
                                "phone": "", 
                                "whatsapp": "",
                                "natchathiram": str(row[col_star]),
                                "dob": format_date_for_db(row[col_dob]),
                                "wedding_date": format_date_for_db(row[col_wed]),
                                "yearly_pooja_date": format_date_for_db(row[col_pj])
                            }
                            run_supabase_insert("members", m_data)
                    
                    st.success(f"Successfully uploaded {success_count} families!")
                    st.rerun()
            except ImportError:
                st.error("Missing optional dependency 'openpyxl'. Please ensure it is added to your requirements.txt file.")
            except Exception as e:
                st.error(f"Failed to process file: {e}")

elif st.session_state.current_page == "Search":
    page_header()
    render_navigation_bar()
    st.header("Search & Manage Devotees")
    
    fams = get_data("families")
    mems = get_data("members")
    
    if fams.empty:
        st.info("No devotees enrolled yet.")
    else:
        all_rows = []
        for _, f in fams.iterrows():
            all_rows.append({
                "fid": f['id'], "pid": f['id'], "ish": True,
                "Family Head Name": f['head_name'], "Address": f['address'],
                "Mobile No": f['phone'], "WhatsApp No": f['whatsapp'],
                "Family Member Name": f['head_name'], "Relationship": "Head",
                "DOB": f['dob'], "Natchathiram": f['natchathiram'],
                "Wedding Day": f['wedding_date'], "Yearly Pooja": f['yearly_pooja_date']
            })
            if not mems.empty:
                f_mems = mems[mems['family_id'] == f['id']]
                for _, m in f_mems.iterrows():
                    all_rows.append({
                        "fid": f['id'], "pid": m['id'], "ish": False,
                        "Family Head Name": f['head_name'], "Address": f['address'],
                        "Mobile No": m['phone'] if m['phone'] else f['phone'], 
                        "WhatsApp No": m['whatsapp'] if m['whatsapp'] else f['whatsapp'],
                        "Family Member Name": m['member_name'], "Relationship": m['relationship'],
                        "DOB": m['dob'], "Natchathiram": m['natchathiram'],
                        "Wedding Day": m['wedding_date'], "Yearly Pooja": m['yearly_pooja_date']
                    })
        
        full_df = pd.DataFrame(all_rows)
        search_term = st.text_input("Search by Name, Mobile, or Address")
        if search_term:
            s_mask = (full_df['Family Head Name'].str.contains(search_term, case=False, na=False) |
                      full_df['Family Member Name'].str.contains(search_term, case=False, na=False) |
                      full_df['Mobile No'].str.contains(search_term, na=False) |
                      full_df['Address'].str.contains(search_term, case=False, na=False))
            full_df = full_df[s_mask]
        
        if full_df.empty:
            st.warning("No records found.")
        else:
            display_df = full_df.copy()
            display_df['DOB'] = display_df['DOB'].apply(format_date_for_ui)
            display_df['Wedding Day'] = display_df['Wedding Day'].apply(format_date_for_ui)
            display_df['Yearly Pooja'] = display_df['Yearly Pooja'].apply(format_date_for_ui)
            display_df = display_df.reset_index(drop=True)
            display_df.insert(0, 'Sl.No', display_df.index + 1)
            final_cols = ['Sl.No', 'Family Head Name', 'Address', 'Mobile No', 'WhatsApp No', 
                         'Family Member Name', 'Relationship', 'DOB', 'Natchathiram', 
                         'Wedding Day', 'Yearly Pooja']
            st.dataframe(display_df[final_cols], hide_index=True, use_container_width=True)
            
            st.divider()
            st.subheader("⚙️ Account Management")
            m_tab1, m_tab2, m_tab3 = st.tabs(["Edit Profiles", "Delete Profiles", "Account View"])
            
            with m_tab1:
                st.write("✏️ **Edit Profiles**")
                edit_type = st.radio("What would you like to edit?", ["Family Head", "Family Member"], horizontal=True)
                
                if edit_type == "Family Head":
                    unique_heads = full_df.drop_duplicates('fid')
                    h_dict = {f"{r['Family Head Name']} (ID: {r['fid']})": r['fid'] for _, r in unique_heads.iterrows()}
                    h_sel = st.selectbox("Select Head to Edit", list(h_dict.keys()))
                    curr_head = fams[fams['id'] == h_dict[h_sel]].iloc[0]
                    
                    with st.form("edit_head_form"):
                        e_h_name = st.text_input("Name", value=curr_head['head_name'])
                        e_h_phone = st.text_input("Phone", value=curr_head['phone'])
                        e_h_wa = st.text_input("WhatsApp", value=curr_head['whatsapp'])
                        e_h_addr = st.text_area("Address", value=curr_head['address'])
                        e_h_star = st.selectbox("Star", NATCHATHIRAM_OPTIONS, index=NATCHATHIRAM_OPTIONS.index(curr_head['natchathiram']) if curr_head['natchathiram'] in NATCHATHIRAM_OPTIONS else 0)
                        
                        e_h_dob = st.date_input("DOB", value=safe_date_convert(curr_head['dob']), min_value=MIN_DATE)
                        e_h_wed = st.date_input("Wedding Day", value=safe_date_convert(curr_head['wedding_date']), min_value=MIN_DATE)
                        e_h_pj = st.date_input("Yearly Pooja", value=safe_date_convert(curr_head['yearly_pooja_date']))
                        
                        if st.form_submit_button("Update Head Profile"):
                            up_data = {
                                "head_name": e_h_name, "phone": e_h_phone, "whatsapp": e_h_wa, 
                                "address": e_h_addr, "natchathiram": e_h_star,
                                "dob": format_date_for_db(e_h_dob), "wedding_date": format_date_for_db(e_h_wed),
                                "yearly_pooja_date": format_date_for_db(e_h_pj)
                            }
                            run_supabase_update("families", up_data, curr_head['id'])
                            st.success("Head Profile Updated!")
                            st.rerun()

                else:
                    if not mems.empty:
                        m_dict = {f"{r['Family Member Name']} ({r['Relationship']}) - Head: {r['Family Head Name']}": r['pid'] for _, r in full_df[full_df['ish'] == False].iterrows()}
                        m_sel = st.selectbox("Select Member to Edit", list(m_dict.keys()))
                        curr_mem = mems[mems['id'] == m_dict[m_sel]].iloc[0]
                        
                        with st.form("edit_mem_form"):
                            e_m_name = st.text_input("Name", value=curr_mem['member_name'])
                            e_m_rel = st.selectbox("Relationship", RELATIONSHIP_OPTIONS, index=RELATIONSHIP_OPTIONS.index(curr_mem['relationship']) if curr_mem['relationship'] in RELATIONSHIP_OPTIONS else 0)
                            e_m_phone = st.text_input("Phone", value=curr_mem['phone'] if curr_mem['phone'] else "")
                            e_m_wa = st.text_input("WhatsApp", value=curr_mem['whatsapp'] if curr_mem['whatsapp'] else "")
                            e_m_star = st.selectbox("Star", NATCHATHIRAM_OPTIONS, index=NATCHATHIRAM_OPTIONS.index(curr_mem['natchathiram']) if curr_mem['natchathiram'] in NATCHATHIRAM_OPTIONS else 0)
                            
                            e_m_dob = st.date_input("DOB", value=safe_date_convert(curr_mem['dob']), min_value=MIN_DATE)
                            e_m_wed = st.date_input("Wedding Day", value=safe_date_convert(curr_mem['wedding_date']), min_value=MIN_DATE)
                            e_m_pj = st.date_input("Yearly Pooja", value=safe_date_convert(curr_mem['yearly_pooja_date']))
                            
                            if st.form_submit_button("Update Member Profile"):
                                up_m_data = {
                                    "member_name": e_m_name, "relationship": e_m_rel, "phone": e_m_phone, 
                                    "whatsapp": e_m_wa, "natchathiram": e_m_star,
                                    "dob": format_date_for_db(e_m_dob), "wedding_date": format_date_for_db(e_m_wed),
                                    "yearly_pooja_date": format_date_for_db(e_m_pj)
                                }
                                run_supabase_update("members", up_m_data, curr_mem['id'])
                                st.success("Member Profile Updated!")
                                st.rerun()
                    else: st.info("No members to edit.")

            with m_tab2:
                st.write("🗑️ **Delete Profiles**")
                del_col1, del_col2 = st.columns(2)
                with del_col1:
                    del_m_list = full_df[full_df['ish'] == False]
                    if not del_m_list.empty:
                        d_m_dict = {f"{r['Family Member Name']} ({r['Relationship']})": r['pid'] for _, r in del_m_list.iterrows()}
                        d_m_sel = st.selectbox("Member to Delete", list(d_m_dict.keys()))
                        if st.button("Delete Member"):
                            run_supabase_delete("members", d_m_dict[d_m_sel])
                            st.rerun()
                with del_col2:
                    d_f_dict = {f"{r['Family Head Name']} (ID: {r['fid']})": r['fid'] for _, r in full_df.drop_duplicates('fid').iterrows()}
                    d_f_sel = st.selectbox("Family to Delete", list(d_f_dict.keys()))
                    if st.button("Delete Entire Family"):
                        supabase.table("members").delete().eq("family_id", d_f_dict[d_f_sel]).execute()
                        run_supabase_delete("families", d_f_dict[d_f_sel])
                        st.rerun()

            with m_tab3:
                st.write("🔍 **View Profile Details**")
                v_heads = full_df.drop_duplicates('fid')
                v_f_dict = {f"{r['Family Head Name']}": r['fid'] for _, r in v_heads.iterrows()}
                v_sel = st.selectbox("Select Family to View", list(v_f_dict.keys()))
                v_row = fams[fams['id'] == v_f_dict[v_sel]].iloc[0]
                
                v_c1, v_c2 = st.columns([1, 3])
                with v_c1:
                    img = base64_to_image(v_row['photo'])
                    if img: st.image(img, width=150)
                    else: st.markdown("👤")
                with v_c2:
                    st.write(f"**Name:** {v_row['head_name']}")
                    st.write(f"**Contact:** {v_row['phone']} / WhatsApp: {v_row['whatsapp']}")
                    st.write(f"**Address:** {v_row['address']}")

elif st.session_state.current_page == "Billing":
    page_header()
    render_navigation_bar()
    st.header("Billing Desk")
    bill_tab1, bill_tab2 = st.tabs(["New Bill", "Bill History"])
    
    with bill_tab1:
        billing_mode = st.radio("Billing Mode", ["Enrolled Devotee", "Guest Devotee"], horizontal=True)
        col1, col2 = st.columns([2, 1])
        with col1:
            selected_name, selected_wa, selected_id, selected_address = "", "", 0, ""
            if billing_mode == "Enrolled Devotee":
                f_data = get_data("families")
                m_data = get_data("members")
                if not f_data.empty:
                    opts = {}
                    for _, f in f_data.iterrows():
                        key = f"{f['head_name']} (Head)"
                        opts[key] = {"id": f['id'], "name": f['head_name'], "address": f['address'], "wa": f['whatsapp'] if f['whatsapp'] else f['phone']}
                    for _, m in m_data.iterrows():
                        key = f"{m['member_name']} ({m['relationship']})"
                        h_addr = f_data[f_data['id'] == m['family_id']]['address'].values[0] if not f_data[f_data['id'] == m['family_id']].empty else "N/A"
                        opts[key] = {"id": m['family_id'], "name": m['member_name'], "address": h_addr, "wa": m['whatsapp'] if m['whatsapp'] else (m['phone'] if m['phone'] else opts.get(f"{f_data[f_data['id'] == m['family_id']]['head_name'].values[0]} (Head)", {}).get("wa", ""))}
                    sel_k = st.selectbox("Select Devotee", list(opts.keys()))
                    d_obj = opts[sel_k]
                    selected_name, selected_id, selected_address, selected_wa = d_obj['name'], d_obj['id'], d_obj['address'], d_obj['wa']
                else: st.warning("Enroll devotees first.")
            else:
                selected_name = st.text_input("Guest Name *")
                selected_address = st.text_area("Guest Address")
                selected_wa = st.text_input("Guest WhatsApp No.")

            servs = get_data("services")
            if not servs.empty:
                s_dict = {r['service_name']: r for _, r in servs.iterrows()}
                sel_s = st.selectbox("Select Service", list(s_dict.keys()))
                srv = s_dict[sel_s]
                man_no = st.text_input("Manual Bill No."); book_no = st.text_input("Bill Book No.")
                if st.button("Generate Receipt"):
                    if billing_mode == "Guest Devotee" and not selected_name: st.error("Enter Name.")
                    else:
                        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        payload = {"family_id": selected_id, "service_id": srv['id'], "amount": srv['price'], "date": now, "manual_bill_no": man_no, "bill_book_no": book_no, "guest_name": selected_name if selected_id == 0 else "", "guest_address": selected_address, "guest_whatsapp": selected_wa}
                        res = run_supabase_insert("transactions", payload)
                        if res and res.data:
                            last_id = res.data[0]['id']
                            pdf = generate_pdf(last_id, selected_name, {selected_address}, sel_s, srv['price'], now, man_no, book_no)
                            st.success("Bill Generated!")
                            c_dl1, c_dl2 = st.columns(2)
                            with c_dl1: st.download_button("📥 PDF Receipt", pdf, f"Rec_{last_id}.pdf")
                            if selected_wa:
                                msg = f"🙏 *{TEMPLE_NAME_FULL}*\nNamaste *{selected_name}*,\nReceipt: *#{last_id}*\nManual: *{man_no}*\nBook: *{book_no}*\nService: *{sel_s}*\nAmount: *₹{srv['price']:,.2f}*\nDate: *{now}*"
                                encoded = msg.replace(' ', '%20').replace('\n', '%0A').replace('#', '%23').replace('*', '%2A')
                                with c_dl2: st.link_button("📲 Send WhatsApp", f"https://wa.me/{selected_wa}?text={encoded}")
            else: st.warning("Add services in Settings.")
        with col2:
            st.subheader("Last 10 Bills")
            tr_history = get_data("transactions")
            if not tr_history.empty:
                st.dataframe(tr_history[['id', 'amount', 'date']].sort_values('id', ascending=False).head(10), hide_index=True)

    with bill_tab2:
        st.subheader("Manage Bills")
        tr_df = get_data("transactions")
        sv_df = get_data("services")
        if not tr_df.empty:
            for _, r in tr_df.sort_values('id', ascending=False).iterrows():
                with st.container():
                    col_b1, col_b2 = st.columns([4, 1])
                    s_nm = sv_df[sv_df['id'] == r['service_id']]['service_name'].values[0] if not sv_df[sv_df['id'] == r['service_id']].empty else "Service"
                    with col_b1: st.write(f"**Receipt #{r['id']}** | {r['guest_name'] if r['family_id']==0 else 'Enrolled'} | {s_nm} | ₹{r['amount']}")
                    with col_b2:
                        if st.session_state.role == ADMIN_ROLE:
                            if st.button("🗑️", key=f"del_b_{r['id']}"): 
                                run_supabase_delete("transactions", r['id'])
                                st.rerun()
                st.divider()

elif st.session_state.current_page == "Expenses":
    page_header()
    render_navigation_bar()
    st.header("Expenses")
    
    # Dynamic categories from DB
    cat_df = get_data("expense_categories")
    categories = cat_df['category_name'].tolist() if not cat_df.empty else DEFAULT_EXPENSE_TYPES

    tab_add, tab_view = st.tabs(["Record Expense", "Expense History"])
    
    with tab_add:
        with st.form("exp_f"):
            en = st.text_input("Title *"); et = st.selectbox("Type", categories)
            ea = st.number_input("Amount", min_value=0.0); ed = st.date_input("Date", value=date.today())
            if st.form_submit_button("Record Expense"):
                if en and ea > 0:
                    run_supabase_insert("users_expenses", {"expense_name": en, "expense_type": et, "amount": ea, "payment_date": str(ed), "status": "Paid"})
                    st.success("Recorded.")
                    st.rerun()

    with tab_view:
        exp_df = get_data("users_expenses")
        if exp_df.empty:
            st.info("No expenses recorded yet.")
        else:
            st.dataframe(exp_df.sort_values('payment_date', ascending=False), use_container_width=True)
            
            st.divider()
            st.subheader("Manage Expenses")
            for _, row in exp_df.sort_values('payment_date', ascending=False).iterrows():
                with st.container():
                    ec1, ec2 = st.columns([4, 1])
                    with ec1:
                        st.write(f"**{row['expense_name']}** | {row['expense_type']} | ₹{float(row['amount']):,.2f}")
                        st.write(f"Date: {row['payment_date']} | Status: {row['status']}")
                    with ec2:
                        if st.session_state.role == ADMIN_ROLE:
                            if st.button("🗑️", key=f"del_exp_{row['id']}"):
                                run_supabase_delete("users_expenses", row['id'])
                                st.rerun()
                st.divider()

elif st.session_state.current_page == "Reports":
    page_header()
    render_navigation_bar()
    st.header("Financial Reports")
    
    # Report Filtering Section
    st.markdown("### 📊 Select Report Parameters")
    report_mode = st.radio("Report Period:", ["Daily", "Weekly", "Monthly", "Custom Date Range"], horizontal=True)
    
    today = date.today()
    start_d, end_d = today, today
    
    if report_mode == "Daily":
        start_d = st.date_input("Select Date", value=today)
        end_d = start_d
    elif report_mode == "Weekly":
        # Streamlit doesn't have a direct week picker, so we pick a day and find the week
        ref_date = st.date_input("Select a day in the target week", value=today)
        start_d = ref_date - timedelta(days=ref_date.weekday()) # Monday
        end_d = start_d + timedelta(days=6) # Sunday
        st.info(f"Report for week: {start_d.strftime('%d/%m/%Y')} to {end_d.strftime('%d/%m/%Y')}")
    elif report_mode == "Monthly":
        c_m, c_y = st.columns(2)
        with c_m: sel_month = st.selectbox("Month", list(calendar.month_name)[1:], index=today.month-1)
        with c_y: sel_year = st.number_input("Year", min_value=2020, max_value=2050, value=today.year)
        month_idx = list(calendar.month_name).index(sel_month)
        start_d = date(sel_year, month_idx, 1)
        last_day = calendar.monthrange(sel_year, month_idx)[1]
        end_d = date(sel_year, month_idx, last_day)
    else:
        cs1, cs2 = st.columns(2)
        with cs1: start_d = st.date_input("Start Date", value=today - timedelta(days=30))
        with cs2: end_d = st.date_input("End Date", value=today)

    # Fetch Data
    df_trans = get_data("transactions")
    df_exp = get_data("users_expenses")
    df_serv = get_data("services")

    # Filter Data
    if not df_trans.empty:
        df_trans['dt'] = pd.to_datetime(df_trans['date']).dt.date
        df_trans = df_trans[(df_trans['dt'] >= start_d) & (df_trans['dt'] <= end_d)]
    
    if not df_exp.empty:
        df_exp['dt'] = pd.to_datetime(df_exp['payment_date']).dt.date
        df_exp = df_exp[(df_exp['dt'] >= start_d) & (df_exp['dt'] <= end_d)]

    # Metrics
    t_inc = df_trans['amount'].sum() if not df_trans.empty else 0
    t_exp = df_exp['amount'].sum() if not df_exp.empty else 0
    t_net = t_inc - t_exp

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("Period Income", f"₹ {t_inc:,.2f}")
    c2.metric("Period Expenses", f"₹ {t_exp:,.2f}")
    c3.metric("Net Profit", f"₹ {t_net:,.2f}")

    # Tables & Downloads
    tab_inc, tab_exp = st.tabs(["Income Details", "Expense Details"])
    
    with tab_inc:
        if not df_trans.empty:
            # Map service names for readable report
            if not df_serv.empty:
                s_map = dict(zip(df_serv['id'], df_serv['service_name']))
                df_trans['Service'] = df_trans['service_id'].map(s_map)
            
            report_inc_df = df_trans[['dt', 'guest_name', 'Service', 'amount', 'manual_bill_no', 'bill_book_no']].rename(columns={'dt': 'Date', 'guest_name': 'Devotee', 'amount': 'Amount', 'manual_bill_no': 'Manual No', 'bill_book_no': 'Book No'})
            st.dataframe(report_inc_df, use_container_width=True, hide_index=True)
        else:
            st.info("No income records found for this period.")

    with tab_exp:
        if not df_exp.empty:
            report_exp_df = df_exp[['dt', 'expense_name', 'expense_type', 'amount', 'payment_method', 'status']].rename(columns={'dt': 'Date', 'expense_name': 'Title', 'expense_type': 'Category', 'amount': 'Amount', 'payment_method': 'Method'})
            st.dataframe(report_exp_df, use_container_width=True, hide_index=True)
        else:
            st.info("No expense records found for this period.")

    # Combined Downloads
    if t_inc > 0 or t_exp > 0:
        st.divider()
        st.markdown("### 📥 Download Reports")
        d_col1, d_col2 = st.columns(2)
        
        with d_col1:
            title_str = f"Financial Report: {start_d.strftime('%d/%m/%Y')} - {end_d.strftime('%d/%m/%Y')}"
            inc_data = report_inc_df if not df_trans.empty else pd.DataFrame()
            exp_data = report_exp_df if not df_exp.empty else pd.DataFrame()
            pdf_data = generate_financial_pdf(inc_data, exp_data, title_str, t_inc, t_exp, t_net)
            st.download_button("📂 Download PDF Report", pdf_data, f"Report_{start_d}.pdf", "application/pdf", use_container_width=True)
            
        with d_col2:
            # Simple Excel export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                if not df_trans.empty: report_inc_df.to_excel(writer, sheet_name='Income', index=False)
                if not df_exp.empty: report_exp_df.to_excel(writer, sheet_name='Expenses', index=False)
            st.download_button("📊 Download Excel Report", output.getvalue(), f"Report_{start_d}.xlsx", use_container_width=True)

elif st.session_state.current_page == "Assets":
    page_header()
    render_navigation_bar()
    st.header("Assets")
    with st.form("ast_f"):
        an = st.text_input("Asset Name"); av = st.number_input("Value")
        if st.form_submit_button("Save"): run_supabase_insert("assets", {"asset_name": an, "value": av})
    st.dataframe(get_data("assets"), use_container_width=True)

elif st.session_state.current_page == "Settings":
    page_header()
    render_navigation_bar()
    st.header("Settings")
    
    s_tab1, s_tab2 = st.tabs(["Service Settings", "Expense Type Settings"])
    
    with s_tab1:
        with st.form("add_svc"):
            sn = st.text_input("Service Name"); sp = st.number_input("Price", min_value=0.0)
            if st.form_submit_button("Add Service"):
                run_supabase_insert("services", {"service_name": sn, "price": sp})
                st.rerun()
        
        st.write("Current Services:")
        serv_list = get_data("services")
        if not serv_list.empty:
            for _, srv_row in serv_list.iterrows():
                sc1, sc2 = st.columns([4, 1])
                sc1.write(f"{srv_row['service_name']} - ₹{srv_row['price']}")
                if sc2.button("🗑️", key=f"del_svc_{srv_row['id']}"):
                    run_supabase_delete("services", srv_row['id'])
                    st.rerun()

    with s_tab2:
        with st.form("add_exp_type"):
            new_cat = st.text_input("New Expense Category Name")
            if st.form_submit_button("Add Expense Type"):
                if new_cat:
                    run_supabase_insert("expense_categories", {"category_name": new_cat})
                    st.rerun()
        
        st.write("Current Expense Categories:")
        cat_list = get_data("expense_categories")
        if not cat_list.empty:
            for _, cat_row in cat_list.iterrows():
                ec1, ec2 = st.columns([4, 1])
                ec1.write(cat_row['category_name'])
                if ec2.button("🗑️", key=f"del_cat_{cat_row['id']}"):
                    run_supabase_delete("expense_categories", cat_row['id'])
                    st.rerun()

elif st.session_state.current_page == "Users":
    page_header()
    render_navigation_bar()
    st.header("Users")
    if st.session_state.role == ADMIN_ROLE:
        # RIGHTS MANAGEMENT UI
        st.subheader("Manage User Rights")
        users_df = get_data("users", "id, username, role, rights")
        if not users_df.empty:
            # Filter out current admin to avoid self-lockout
            other_users = users_df[users_df['username'] != st.session_state.username]
            if not other_users.empty:
                user_to_mod = st.selectbox("Select User to Assign Rights", other_users['username'].tolist())
                selected_user_row = other_users[other_users['username'] == user_to_mod].iloc[0]
                
                # Get current rights or default to Home
                current_rights = selected_user_row['rights'].split(',') if selected_user_row['rights'] else ["Home Dashboard"]
                
                new_rights = st.multiselect(
                    "Assign Menu Access (Rights)", 
                    options=ALL_MENU_KEYS, 
                    default=[r for r in current_rights if r in ALL_MENU_KEYS]
                )
                
                if st.button("Update User Rights"):
                    rights_str = ",".join(new_rights) if new_rights else "Home Dashboard"
                    run_supabase_update("users", {"rights": rights_str}, selected_user_row['id'])
                    st.success(f"Rights updated for {user_to_mod}!")
                    st.rerun()
            else:
                st.info("No other users found to manage.")

        st.divider()
        st.subheader("Create New User")
        with st.form("u_f"):
            un = st.text_input("User"); up = st.text_input("Pass", type="password")
            if st.form_submit_button("Create"):
                # New users get Home Dashboard by default
                run_supabase_insert("users", {"username": un, "password_hash": hash_password(up), "role": "user", "rights": "Home Dashboard"})
                st.rerun()
        st.dataframe(get_data("users", "id, username, role, rights"), use_container_width=True)

render_footer()
