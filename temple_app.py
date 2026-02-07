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
    if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nat":
        return None
    try:
        d_obj = pd.to_datetime(val)
        return d_obj.strftime('%Y-%m-%d')
    except:
        return None

def format_date_for_ui(val):
    if not val or str(val).lower() in ["none", "nat", ""]:
        return ""
    try:
        d_obj = pd.to_datetime(val)
        return d_obj.strftime('%d/%m/%Y')
    except:
        return str(val)

def safe_date_convert(val):
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
    
    story.append(Paragraph("Income Details", styles['h4']))
    if not income_df.empty:
        inc_pdf_df = income_df.copy()
        if 'Date' in inc_pdf_df.columns: inc_pdf_df['Date'] = inc_pdf_df['Date'].astype(str)
        inc_data = [inc_pdf_df.columns.tolist()] + inc_pdf_df.values.tolist()
        t_inc_table = Table(inc_data, hAlign='LEFT', repeatRows=1)
        t_inc_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#800000')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.gold),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        story.append(t_inc_table)
    
    story.append(Spacer(1, 24))
    story.append(Paragraph("Expense Details", styles['h4']))
    if not expense_df.empty:
        exp_pdf_df = expense_df.copy()
        if 'Date' in exp_pdf_df.columns: exp_pdf_df['Date'] = exp_pdf_df['Date'].astype(str)
        exp_data = [exp_pdf_df.columns.tolist()] + exp_pdf_df.values.tolist()
        t_exp_table = Table(exp_data, hAlign='LEFT', repeatRows=1)
        t_exp_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#800000')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.gold),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        story.append(t_exp_table)

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
            user_rights = user_data.get('rights', 'Home Dashboard').split(',')
            return True, user_data['role'], user_rights
    return False, None, None

# --- UI COMPONENTS ---
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
        }}
        </style>
        """
    st.markdown(bg_style, unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown("<div style='height: 15vh;'></div>", unsafe_allow_html=True)
        st.markdown("<h1 style='font-size: 60px; text-align: center; color: #800000;'>🕉️</h1>", unsafe_allow_html=True)
        st.markdown("<p style='color: #800000; font-weight: bold; text-align: center;'>Amme Narayana... Devi Narayana...</p>", unsafe_allow_html=True)
        st.markdown("""<div style='background-color: #800000; color: #FFD700; padding: 15px; text-align: center; font-weight: bold; font-size: 24px; border: 2px solid #FFD700;'>Temple Management System</div>""", unsafe_allow_html=True)
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
            height: 3.5em; width: 100%; transition: all 0.3s ease;
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

def render_news_ticker():
    today = date.today()
    today_md = today.strftime('%m-%d')
    ticker_items = []
    df_fam = get_data("families")
    if not df_fam.empty:
        b_df = df_fam[df_fam['dob'].astype(str).str.contains(today_md, na=False)]
        for _, r in b_df.iterrows(): ticker_items.append(f"🎂 Birthday: {r['head_name']}!")
        p_df = df_fam[df_fam['yearly_pooja_date'].astype(str).str.contains(today_md, na=False)]
        for _, r in p_df.iterrows(): ticker_items.append(f"🙏 Pooja Reminder: {r['head_name']}!")
    scrolling_text = " | ".join(ticker_items) if ticker_items else "✨ Welcome to Sree Bhadreshwari Amman Temple Management System. ✨"
    st.markdown(f"""
        <style>
        .ticker-wrap {{ width: 100%; overflow: hidden; background-color: #800000; padding: 10px 0; border: 2px solid #FFD700; margin-bottom: 20px; }}
        .ticker {{ display: inline-block; white-space: nowrap; animation: marquee 30s linear infinite; color: #FFD700; font-weight: bold; }}
        @keyframes marquee {{ 0% {{ transform: translateX(100%); }} 100% {{ transform: translateX(-100%); }} }}
        </style>
        <div class="ticker-wrap"><div class="ticker">{scrolling_text}</div></div>
    """, unsafe_allow_html=True)

# --- APP INITIALIZATION ---
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "Home Dashboard"
if 'new_family_id' not in st.session_state: st.session_state.new_family_id = None

if not st.session_state.logged_in:
    login_page()
    render_footer() 
    st.stop()

# --- MODULES ---

if st.session_state.current_page == "Home Dashboard":
    page_header()
    render_navigation_bar()
    render_news_ticker()
    st.title(f"Welcome, {st.session_state.username.title()}")
    today = date.today()
    df_trans = get_data("transactions")
    df_fam = get_data("families")
    col1, col2 = st.columns(2)
    col1.metric("Total Devotees", len(df_fam))
    if not df_trans.empty:
        df_trans['dt'] = pd.to_datetime(df_trans['date']).dt.date
        daily_inc = df_trans[df_trans['dt'] == today]['amount'].sum()
        col2.metric("Today's Collection", f"₹ {daily_inc:,.2f}")

elif st.session_state.current_page == "Enroll":
    page_header(); render_navigation_bar(); st.header("Devotee Enrollment")
    tab1, tab2 = st.tabs(["Manual Entry", "Bulk Upload"])
    with tab1:
        with st.form("head_f"):
            c1, c2 = st.columns(2)
            hn = c1.text_input("Head Name *"); hp = c1.text_input("Phone *"); hwa = c1.text_input("WhatsApp")
            ha = c2.text_area("Address"); hdb = c2.date_input("DOB", value=None, min_value=MIN_DATE)
            hnatch = c1.selectbox("Star", [""] + NATCHATHIRAM_OPTIONS)
            hpj = c2.date_input("Yearly Pooja Date", value=None)
            if st.form_submit_button("Save Family Head"):
                if hn and hp:
                    res = run_supabase_insert("families", {"head_name": hn, "phone": hp, "whatsapp": hwa, "address": ha, "dob": format_date_for_db(hdb), "natchathiram": hnatch, "yearly_pooja_date": format_date_for_db(hpj)})
                    if res: st.success("Saved!"); st.rerun()

elif st.session_state.current_page == "Search":
    page_header(); render_navigation_bar(); st.header("Search Devotees")
    fams = get_data("families")
    if not fams.empty:
        sq = st.text_input("Search Name/Phone/Address")
        if sq: fams = fams[fams['head_name'].str.contains(sq, case=False) | fams['phone'].str.contains(sq)]
        st.dataframe(fams, use_container_width=True)

elif st.session_state.current_page == "Billing":
    page_header(); render_navigation_bar(); st.header("Billing Desk")
    mode = st.radio("Mode", ["Enrolled", "Guest"], horizontal=True)
    with st.form("bill_f"):
        name = st.text_input("Name"); addr = st.text_area("Address")
        servs = get_data("services")
        s_dict = {r['service_name']: r for _, r in servs.iterrows()} if not servs.empty else {}
        sel_s = st.selectbox("Service", list(s_dict.keys()))
        amt = s_dict[sel_s]['price'] if sel_s else 0
        st.write(f"**Amount: ₹ {amt}**")
        if st.form_submit_button("Generate Bill"):
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            sid = s_dict[sel_s]['id']
            res = run_supabase_insert("transactions", {"guest_name": name, "guest_address": addr, "amount": amt, "service_id": sid, "date": now, "family_id": 0})
            if res: st.success("Bill Generated!"); st.rerun()

elif st.session_state.current_page == "Expenses":
    page_header(); render_navigation_bar(); st.header("Expenses")
    with st.form("exp_f"):
        en = st.text_input("Title"); et = st.selectbox("Type", DEFAULT_EXPENSE_TYPES); ea = st.number_input("Amount")
        if st.form_submit_button("Record"):
            run_supabase_insert("users_expenses", {"expense_name": en, "expense_type": et, "amount": ea, "payment_date": str(date.today()), "status": "Paid"})
            st.rerun()

elif st.session_state.current_page == "Reports":
    page_header()
    render_navigation_bar()
    st.header("Financial & Service-wise Reports")
    
    # --- DATE FILTERING SECTION ---
    report_mode = st.radio("Report Period:", ["Daily", "Weekly", "Monthly", "Custom Date Range"], horizontal=True)
    today = date.today()
    start_d, end_d = today, today

    if report_mode == "Daily":
        start_d = st.date_input("Select Date", value=today)
        end_d = start_d
    elif report_mode == "Weekly":
        ref_date = st.date_input("Select a day in the target week", value=today)
        start_d = ref_date - timedelta(days=ref_date.weekday())
        end_d = start_d + timedelta(days=6)
    elif report_mode == "Monthly":
        c_m, c_y = st.columns(2)
        sel_month = c_m.selectbox("Month", list(calendar.month_name)[1:], index=today.month-1)
        sel_year = c_y.number_input("Year", min_value=2020, value=today.year)
        month_idx = list(calendar.month_name).index(sel_month)
        start_d = date(sel_year, month_idx, 1)
        end_d = date(sel_year, month_idx, calendar.monthrange(sel_year, month_idx)[1])
    else:
        cs1, cs2 = st.columns(2)
        start_d = cs1.date_input("Start Date", value=today-timedelta(30))
        end_d = cs2.date_input("End Date", value=today)

    # --- DATA FETCHING ---
    df_trans = get_data("transactions")
    df_exp = get_data("users_expenses")
    df_serv = get_data("services")

    if not df_trans.empty: 
        df_trans['dt'] = pd.to_datetime(df_trans['date']).dt.date
        df_trans = df_trans[(df_trans['dt'] >= start_d) & (df_trans['dt'] <= end_d)]
    if not df_exp.empty: 
        df_exp['dt'] = pd.to_datetime(df_exp['payment_date']).dt.date
        df_exp = df_exp[(df_exp['dt'] >= start_d) & (df_exp['dt'] <= end_d)]

    t_inc = df_trans['amount'].sum() if not df_trans.empty else 0
    t_exp = df_exp['amount'].sum() if not df_exp.empty else 0
    t_net = t_inc - t_exp

    st.divider()
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Income", f"₹ {t_inc:,.2f}")
    m2.metric("Total Expenses", f"₹ {t_exp:,.2f}")
    m3.metric("Net Profit", f"₹ {t_net:,.2f}")

    # --- SERVICE-WISE BREAKDOWN ---
    st.markdown("### 🕉️ Service-wise Collection")
    if not df_trans.empty and not df_serv.empty:
        df_merged = df_trans.merge(df_serv, left_on='service_id', right_on='id', suffixes=('', '_master'))
        svc_report = df_merged.groupby('service_name').agg({'amount': 'sum', 'id': 'count'}).reset_index()
        svc_report.columns = ['Service Name', 'Total Amount (₹)', 'Count']
        col_t, col_c = st.columns([1, 1])
        col_t.dataframe(svc_report, use_container_width=True, hide_index=True)
        col_c.bar_chart(data=svc_report, x='Service Name', y='Total Amount (₹)', color="#800000")
    else:
        st.info("No service transactions found for this period.")

    # --- DETAILED LEDGER ---
    st.divider()
    st.subheader("📋 Detailed Ledger")
    ledger_rows = []
    if not df_trans.empty:
        for _, r in df_trans.iterrows():
            s_name = df_serv[df_serv['id'] == r['service_id']]['service_name'].values[0] if not df_serv[df_serv['id'] == r['service_id']].empty else "Income"
            ledger_rows.append({"id": r['id'], "source": "transactions", "Date": r['dt'], "Description": f"{s_name} - {r['guest_name']}", "Income": r['amount'], "Expenses": 0})
    if not df_exp.empty:
        for _, r in df_exp.iterrows():
            ledger_rows.append({"id": r['id'], "source": "users_expenses", "Date": r['dt'], "Description": r['expense_name'], "Income": 0, "Expenses": r['amount']})
    
    if ledger_rows:
        l_df = pd.DataFrame(ledger_rows).sort_values("Date")
        st.dataframe(l_df.drop(columns=['id', 'source']), use_container_width=True, hide_index=True)
        st.download_button("📊 Export to Excel", to_excel(l_df), f"Report_{start_d}.xlsx")

elif st.session_state.current_page == "Settings":
    page_header(); render_navigation_bar(); st.header("Settings")
    with st.form("svc_f"):
        sn = st.text_input("New Service Name"); sp = st.number_input("Price")
        if st.form_submit_button("Add"):
            run_supabase_insert("services", {"service_name": sn, "price": sp})
            st.rerun()

elif st.session_state.current_page == "Users" and st.session_state.role == ADMIN_ROLE:
    page_header(); render_navigation_bar(); st.header("User Management")
    with st.form("user_f"):
        un = st.text_input("Username"); up = st.text_input("Password", type="password")
        if st.form_submit_button("Create User"):
            run_supabase_insert("users", {"username": un, "password_hash": hash_password(up), "role": "user", "rights": "Home Dashboard"})
            st.rerun()

render_footer()
