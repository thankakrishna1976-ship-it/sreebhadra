import streamlit as st
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

# Dropdown Options
RELATIONSHIP_OPTIONS = ['Wife', 'Son', 'Daughter', 'Mother', 'Father', 'Grand Father', 'Grand Mother', 'Guardian', 'Other']
NATCHATHIRAM_OPTIONS = ['Ashwini', 'Bharani', 'Karthigai', 'Rohini', 'Mrigasiram', 'Thiruvathirai', 'Punarpoosam', 'Poosam', 'Ayilyam', 'Magam', 'Poorvam', 'Uthiram', 'Hastham', 'Chithirai', 'Swathi', 'Visakam', 'Anusham', 'Kettai', 'Moolam', 'Pooradam', 'Uthiradam', 'Thiruvonam', 'Avittam', 'Sathayam', 'Poorattathi', 'Uthirattathi', 'Revathi']
EXPENSE_TYPES = ['Pooja Items', 'Maintenance/Repairs', 'Salary/Dakshina', 'Electricity/Water', 'Annadanam/Food', 'Construction', 'Festivals', 'Administrative', 'Other']
PAYMENT_METHODS = ['Cash', 'UPI / GPay', 'Bank Transfer', 'Cheque', 'Card']
PAYMENT_STATUS = ['Paid', 'Pending', 'Partial']

MIN_DATE = date(1940, 1, 1)
MAX_DATE = date(2040, 12, 31)

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
            df.to_excel(writer, index=False, sheet_name='Income_Report')
        return output.getvalue()
    except: return None

def generate_income_pdf(df, title, total_income):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=50, bottomMargin=30)
    styles = getSampleStyleSheet()
    story = [Paragraph(TEMPLE_NAME_FULL, styles['Title']), Paragraph(title, styles['h3']), Spacer(1, 12)]
    story.append(Paragraph(f"<b>Total Income:</b> ₹ {total_income:,.2f}", styles['Normal']))
    story.append(Spacer(1, 12))
    df_pdf = df.copy()
    if 'Amount' in df_pdf.columns:
        df_pdf['Amount'] = df_pdf['Amount'].apply(lambda x: f"₹ {float(x):,.2f}")
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

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def verify_user(username, password):
    res = supabase.table('users').select('*').eq('username', username).execute()
    if res.data:
        user_data = res.data[0]
        if user_data['password_hash'] == hash_password(password):
            return True, user_data['role']
    return False, None

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
            success, role = verify_user(username, password)
            if success:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.role = role
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
    NAV_BAR_PAGES = {
        "Home Dashboard": {"label": "HOME"}, "Enroll": {"label": "ENROLLMENT"},
        "Search": {"label": "SEARCH"}, "Billing": {"label": "BILLING"},
        "Expenses": {"label": "EXPENSES"}, "Reports": {"label": "REPORTS"},
        "Assets": {"label": "ASSETS"}, "Settings": {"label": "SETTINGS"},
    }
    if st.session_state.role == ADMIN_ROLE: NAV_BAR_PAGES["Users"] = {"label": "USERS"}
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
            "head_name": head_name, "dob": str(dob) if dob else None,
            "wedding_date": str(wedding_date) if wedding_date else None,
            "natchathiram": n_star, "address": address, "phone": phone,
            "whatsapp": whatsapp, "photo": photo_str,
            "yearly_pooja_date": str(pooja_date) if pooja_date else None
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
        b_df = df_fam[df_fam['dob'].str.contains(today_md, na=False)]
        for _, r in b_df.iterrows(): ticker_items.append(f"🎂 Happy Birthday to Devotee Head: {r['head_name']}!")
        p_df = df_fam[df_fam['yearly_pooja_date'].str.contains(today_md, na=False)]
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
    t_str = date.today().strftime('%Y-%m-%d')
    y_str = date.today().strftime('%Y')
    df_trans = get_data("transactions")
    df_exp_all = get_data("users_expenses")
    inc_t = df_trans[df_trans['date'].str.startswith(t_str, na=False)]['amount'].sum() if not df_trans.empty else 0
    inc_y = df_trans[df_trans['date'].str.startswith(y_str, na=False)]['amount'].sum() if not df_trans.empty else 0
    exp_t = df_exp_all[df_exp_all['payment_date'] == t_str]['amount'].sum() if not df_exp_all.empty else 0
    exp_y = df_exp_all[df_exp_all['payment_date'].str.startswith(y_str, na=False)]['amount'].sum() if not df_exp_all.empty else 0
    df_fam = get_data("families")
    st.subheader("💰 Income Summary")
    c1, c2 = st.columns(2)
    c1.metric("Today's Income", f"₹ {inc_t:,.2f}")
    c2.metric("Yearly Income", f"₹ {inc_y:,.2f}")
    st.subheader("💸 Expense Summary")
    e1, e2 = st.columns(2)
    e1.metric("Today's Expenses", f"₹ {exp_t:,.2f}")
    e2.metric("Yearly Expenses", f"₹ {exp_y:,.2f}")
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
                with col_nm1: m_name = st.text_input("Member Name")
                with col_nm2: m_rel = st.selectbox("Relationship", [""] + RELATIONSHIP_OPTIONS)
                col_cnt1, col_cnt2 = st.columns(2)
                with col_cnt1: m_phone = st.text_input("Member Mobile No.")
                with col_cnt2: m_wa = st.text_input("Member WhatsApp No.")
                if st.form_submit_button("Add Member"):
                    if m_name:
                        data = {"family_id": st.session_state.new_family_id, "member_name": m_name, "relationship": m_rel, "phone": m_phone, "whatsapp": m_wa}
                        run_supabase_insert("members", data)
                        st.success(f"Member {m_name} added!")
                    else: st.error("Name required.")

    with tab_bulk:
        st.subheader("Bulk Import Devotees")
        st.info("""
            **Excel Format Instructions:**
            1. **family_tag**: Use the same unique text/number for a head and their members (e.g., 'F1', 'F2').
            2. **relationship**: Enter 'Head' for the family head. For others, enter 'Wife', 'Son', etc.
            3. **Other columns**: `name`, `phone`, `whatsapp`, `address`, `dob`, `wedding_date`, `natchathiram`, `yearly_pooja_date`.
        """)
        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
        if uploaded_file:
            try:
                bulk_df = pd.read_excel(uploaded_file).fillna("")
                st.write("Preview of data:")
                st.dataframe(bulk_df.head())
                
                if st.button("🚀 Process Bulk Upload"):
                    success_count = 0
                    for tag, group in bulk_df.groupby('family_tag'):
                        head_mask = group['relationship'].str.lower() == 'head'
                        if not head_mask.any():
                            st.error(f"Error: Family Tag '{tag}' has no 'Head' defined. Skipping.")
                            continue
                        
                        head_row = group[head_mask].iloc[0]
                        head_data = {
                            "head_name": str(head_row['name']),
                            "phone": str(head_row['phone']),
                            "whatsapp": str(head_row['whatsapp']),
                            "address": str(head_row['address']),
                            "natchathiram": str(head_row['natchathiram']),
                            "dob": str(head_row['dob']) if head_row['dob'] else None,
                            "wedding_date": str(head_row['wedding_date']) if head_row['wedding_date'] else None,
                            "yearly_pooja_date": str(head_row['yearly_pooja_date']) if head_row['yearly_pooja_date'] else None
                        }
                        
                        h_res = run_supabase_insert("families", head_data)
                        if h_res and h_res.data:
                            new_family_id = h_res.data[0]['id']
                            success_count += 1
                            members_rows = group[~head_mask]
                            for _, m_row in members_rows.iterrows():
                                m_data = {
                                    "family_id": new_family_id,
                                    "member_name": str(m_row['name']),
                                    "relationship": str(m_row['relationship']),
                                    "phone": str(m_row['phone']),
                                    "whatsapp": str(m_row['whatsapp']),
                                    "natchathiram": str(m_row['natchathiram']),
                                    "dob": str(m_row['dob']) if m_row['dob'] else None,
                                    "wedding_date": str(m_row['wedding_date']) if m_row['wedding_date'] else None,
                                    "yearly_pooja_date": str(m_row['yearly_pooja_date']) if m_row['yearly_pooja_date'] else None
                                }
                                run_supabase_insert("members", m_data)
                    st.success(f"Successfully uploaded {success_count} families!")
            except Exception as e:
                st.error(f"Failed to process file: {e}")

elif st.session_state.current_page == "Search":
    page_header()
    render_navigation_bar()
    st.header("Search & Manage Devotees")
    search_term = st.text_input("Search by Name or Mobile No.")
    df = get_data("families")
    if not df.empty:
        if search_term:
            df = df[df['head_name'].str.contains(search_term, case=False, na=False) | df['phone'].str.contains(search_term, na=False)]
        
        if df.empty:
            st.info("No matching devotees found.")
        else:
            for idx, row in df.iterrows():
                with st.container():
                    c_img, c_detail = st.columns([1, 6])
                    with c_img:
                        img_data = base64_to_image(row['photo'])
                        if img_data: st.image(img_data, width=100)
                        else: st.markdown("👤")
                    with c_detail:
                        st.subheader(row['head_name'])
                        st.write(f"📞 {row['phone']} | ⭐ {row['natchathiram'] or 'N/A'}")
                        with st.expander("⚙️ Manage Account"):
                            st.write("👨‍👩‍👧 **Family Members**")
                            members = get_data("members")
                            if not members.empty:
                                fam_m = members[members['family_id'] == row['id']]
                                if fam_m.empty:
                                    st.write("No members added.")
                                else:
                                    for _, m in fam_m.iterrows():
                                        col_m1, col_m2 = st.columns([4, 1])
                                        col_m1.write(f"- {m['member_name']} ({m['relationship']})")
                                        if col_m2.button("🗑️", key=f"del_mem_{m['id']}"):
                                            run_supabase_delete("members", m['id'])
                                            st.rerun()
                            
                            st.divider()
                            st.write("🗑️ **Danger Zone**")
                            if st.button(f"Delete Account: {row['head_name']}", key=f"del_{row['id']}"):
                                supabase.table("members").delete().eq("family_id", row['id']).execute()
                                run_supabase_delete("families", row['id'])
                                st.rerun()
    else:
        st.info("Enrollment database is empty.")

elif st.session_state.current_page == "Billing":
    page_header()
    render_navigation_bar()
    st.header("Billing Desk")
    billing_mode = st.radio("Billing Mode", ["Enrolled Devotee", "Guest Devotee"], horizontal=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        selected_name = ""; selected_wa = ""; selected_id = 0
        if billing_mode == "Enrolled Devotee":
            fams = get_data("families")
            if not fams.empty:
                fam_list = {f"{r['head_name']} ({r['phone']})": r for _, r in fams.iterrows()}
                sel_f = st.selectbox("Choose Devotee", list(fam_list.keys()))
                dev = fam_list[sel_f]
                selected_name, selected_id = dev['head_name'], dev['id']
                selected_wa = str(dev['whatsapp']).strip() if str(dev['whatsapp']).strip() else str(dev['phone']).strip()
            else:
                st.warning("Enroll a devotee first.")
        else:
            selected_name = st.text_input("Guest Name *")
            selected_wa = st.text_input("Guest WhatsApp No.")

        services = get_data("services")
        if not services.empty:
            serv_dict = {r['service_name']: r for _, r in services.iterrows()}
            sel_s = st.selectbox("Select Service", list(serv_dict.keys()))
            service = serv_dict[sel_s]
            man_no = st.text_input("Manual Bill No."); book_no = st.text_input("Bill Book No.")
            
            if st.button("Generate Receipt"):
                if billing_mode == "Guest Devotee" and not selected_name:
                    st.error("Enter Guest Name.")
                else:
                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    data = {"family_id": selected_id, "service_id": service['id'], "amount": service['price'], "date": now, "manual_bill_no": man_no, "bill_book_no": book_no, "guest_name": selected_name if selected_id == 0 else "", "guest_whatsapp": selected_wa}
                    res = run_supabase_insert("transactions", data)
                    if res and res.data:
                        last_id = res.data[0]['id']
                        pdf = generate_pdf(last_id, selected_name, "", sel_s, service['price'], now, man_no, book_no)
                        st.success("Receipt generated successfully!")
                        col_dl1, col_dl2 = st.columns(2)
                        with col_dl1: st.download_button("📥 Download PDF Receipt", pdf, f"Receipt_{last_id}.pdf", "application/pdf")
                        if selected_wa:
                            msg = f"🙏 *{TEMPLE_NAME_FULL}*\n\nNamaste *{selected_name}*,\nReceipt: *#{last_id}*\nManual No: *{man_no}*\nBook No: *{book_no}*\nService: *{sel_s}*\nAmount: *₹{service['price']:,.2f}*\nDate: *{now}*"
                            clean_wa = "".join(filter(str.isdigit, str(selected_wa)))
                            if len(clean_wa) == 10: clean_wa = "91" + clean_wa
                            encoded_msg = msg.replace(' ', '%20').replace('\n', '%0A').replace('#', '%23').replace('*', '%2A')
                            with col_dl2: st.link_button("📲 Send WhatsApp", f"https://wa.me/{clean_wa}?text={encoded_msg}")
        else:
            st.warning("Please add services in Settings first.")

elif st.session_state.current_page == "Expenses":
    page_header()
    render_navigation_bar()
    st.header("Expense Management")
    tab_add, tab_view = st.tabs(["Record Expense", "Expense History"])
    with tab_add:
        with st.form("add_expense"):
            en = st.text_input("Expense Title *"); et = st.selectbox("Category", EXPENSE_TYPES)
            ea = st.number_input("Amount (₹)", min_value=0.0); ed = st.date_input("Date", value=date.today())
            em = st.selectbox("Payment Method", PAYMENT_METHODS); es = st.selectbox("Status", PAYMENT_STATUS)
            ev = st.text_input("Voucher / Bill No."); desc = st.text_area("Description")
            if st.form_submit_button("Save Expense"):
                if en and ea > 0:
                    run_supabase_insert("users_expenses", {"expense_name": en, "expense_type": et, "amount": ea, "payment_date": str(ed), "payment_method": em, "status": es, "voucher_no": ev, "description": desc})
                    st.success("Expense saved.")
    with tab_view:
        st.dataframe(get_data("users_expenses"), use_container_width=True)

elif st.session_state.current_page == "Reports":
    page_header()
    render_navigation_bar()
    st.header("Reports")
    df_i = get_data("transactions")
    if not df_i.empty:
        st.metric("Total Overall Income", f"₹ {df_i['amount'].sum():,.2f}")
        st.dataframe(df_i, use_container_width=True)
    else:
        st.info("No transaction data available for reports.")

elif st.session_state.current_page == "Assets":
    page_header()
    render_navigation_bar()
    st.header("Asset Management")
    with st.expander("Register Asset"):
        with st.form("asset_f"):
            an = st.text_input("Asset Name *"); ad = st.text_input("Description")
            av = st.number_input("Value", min_value=0.0); aq = st.number_input("Qty", min_value=1)
            if st.form_submit_button("Save"):
                run_supabase_insert("assets", {"asset_name": an, "description": ad, "value": av, "quantity": aq})
    st.dataframe(get_data("assets"), use_container_width=True)

elif st.session_state.current_page == "Settings":
    page_header()
    render_navigation_bar()
    st.header("Settings")
    with st.form("add_svc"):
        sn = st.text_input("Service Name"); sp = st.number_input("Price", min_value=0.0)
        if st.form_submit_button("Add Service"):
            run_supabase_insert("services", {"service_name": sn, "price": sp})
            st.rerun()
    st.table(get_data("services"))

elif st.session_state.current_page == "Users":
    page_header()
    render_navigation_bar()
    st.header("User Management")
    if st.session_state.role == ADMIN_ROLE:
        with st.form("user_f"):
            un = st.text_input("Username"); up = st.text_input("Password", type="password")
            ur = st.selectbox("Role", [USER_ROLE, ADMIN_ROLE])
            if st.form_submit_button("Create"):
                run_supabase_insert("users", {"username": un, "password_hash": hash_password(up), "role": ur})
        st.dataframe(get_data("users", "id, username, role"), use_container_width=True)
    else:
        st.error("Admin access required.")

render_footer()
