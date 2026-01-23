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

# Menu Keys
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
        return supabase.table(table_name).insert(data).execute()
    except Exception as e:
        st.error(f"Database Error: {e}")
        return None

def run_supabase_update(table_name, data, row_id):
    try:
        return supabase.table(table_name).update(data).eq('id', row_id).execute()
    except Exception as e:
        st.error(f"Database Error: {e}")
        return None

def run_supabase_delete(table_name, row_id):
    try:
        return supabase.table(table_name).delete().eq('id', row_id).execute()
    except Exception as e:
        st.error(f"Database Error: {e}")
        return None

def get_data(table_name, select="*"):
    try:
        response = supabase.table(table_name).select(select).execute()
        return pd.DataFrame(response.data)
    except Exception as e:
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
    try: return pd.to_datetime(val).date()
    except: return None

def get_base64_of_bin_file(bin_file):
    try:
        if os.path.exists(bin_file):
            with open(bin_file, 'rb') as f: return base64.b64encode(f.read()).decode()
    except: pass
    return None

def image_to_base64(image_file):
    return base64.b64encode(image_file.getvalue()).decode() if image_file else ""

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
    except: return None

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def verify_user(username, password):
    res = supabase.table('users').select('*').eq('username', username).execute()
    if res.data:
        user = res.data[0]
        if user['password_hash'] == hash_password(password):
            rights = user.get('rights', 'Home Dashboard').split(',')
            return True, user['role'], rights
    return False, None, None

def generate_pdf(receipt_no, devotee_name, devotee_address, service, amount, trans_date, manual_no, book_no):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    y = 780 
    try:
        if os.path.exists(LOGO_PATH):
            img = ImageReader(LOGO_PATH)
            c.drawImage(img, 50, y - 50, width=50, height=50)
    except: pass
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(300, y, TEMPLE_NAME_FULL)
    c.setFont("Helvetica", 10)
    c.drawCentredString(300, y - 15, TRUST_DETAILS)
    c.drawCentredString(300, y - 30, ADDRESS_LINE_1)
    c.drawCentredString(300, y - 45, ADDRESS_LINE_2)
    c.line(50, y - 60, 550, y - 60)
    
    dy = y - 90
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, dy, f"RECEIPT No: #{receipt_no}")
    c.drawString(400, dy, f"DATE: {trans_date}")
    c.setFont("Helvetica", 10)
    c.drawString(50, dy - 20, f"Manual Bill No: {manual_no if manual_no else 'N/A'}")
    c.drawString(250, dy - 20, f"Bill Book No: {book_no if book_no else 'N/A'}")
    
    c.setFont("Helvetica", 12)
    c.drawString(50, dy - 50, f"Devotee Name: {devotee_name}")
    c.setFont("Helvetica", 10)
    c.drawString(50, dy - 70, f"Address: {str(devotee_address)[:70]}")
    
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, dy - 100, f"Seva / Pooja: {service}")
    
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, dy - 140, "AMOUNT PAID:")
    c.setFillColorRGB(0.5, 0, 0) 
    c.drawString(180, dy - 140, f"Rs. {float(amount):,.2f}/-")
    
    c.setFillColorRGB(0, 0, 0) 
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(50, dy - 240, "Thank you for your offering. May the blessings of Sree Bhadreshwari Amman be upon you.")
    c.drawString(400, dy - 290, "Authorized Signature")
    c.save()
    buffer.seek(0)
    return buffer

# --- LOGIN ---
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "Home Dashboard"

if not st.session_state.logged_in:
    bg_img_b64 = get_base64_of_bin_file(BACKGROUND_PATH)
    st.markdown(f"""<style>.stApp {{ background-image: url('data:image/jpg;base64,{bg_img_b64}'); background-size: cover; background-color: #bf953f; }}</style>""", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown("<h1 style='text-align: center; color: #800000; padding-top: 15vh;'>🕉️ Staff Login</h1>", unsafe_allow_html=True)
        u_in = st.text_input("Username")
        p_in = st.text_input("Password", type="password")
        if st.button("SIGN IN", use_container_width=True):
            ok, role, rights = verify_user(u_in, p_in)
            if ok:
                st.session_state.update({"logged_in": True, "username": u_in, "role": role, "rights": rights})
                st.rerun()
            else: st.error("Invalid Username or Password.")
    st.stop()

# --- MAIN UI ---
def render_footer():
    st.markdown("""<div style='position: fixed; left: 0; bottom: 0; width: 100%; background-color: #800000; color: #FFD700; text-align: center; padding: 10px 0; font-size: 14px; font-weight: bold; border-top: 2px solid #FFD700; z-index: 999;'>Developed By : Sai Dharshini Info Solution</div>""", unsafe_allow_html=True)

def page_header():
    col_img, col_title = st.columns([1, 8])
    with col_img:
        if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=80)
        else: st.markdown("### 🕉️")
    with col_title:
        st.markdown(f"<h1 style='color: #800000; border-bottom: 2px solid #b38728;'>{TEMPLE_NAME_FULL}</h1>", unsafe_allow_html=True)
    st.divider()

def render_navigation_bar():
    ALL_PAGES = {
        "Home Dashboard": "HOME", "Enroll": "ENROLLMENT", "Search": "SEARCH",
        "Billing": "BILLING", "Expenses": "EXPENSES", "Reports": "REPORTS",
        "Assets": "ASSETS", "Samayavakuppu": "SAMAYAVAKUPPU", "Settings": "SETTINGS"
    }
    nav_keys = list(ALL_PAGES.keys()) if st.session_state.role == ADMIN_ROLE else [k for k in ALL_PAGES.keys() if k in st.session_state.rights]
    if st.session_state.role == ADMIN_ROLE: nav_keys.append("Users")
    
    cols = st.columns(len(nav_keys) + 1)
    for i, k in enumerate(nav_keys):
        label = "USERS" if k == "Users" else ALL_PAGES[k]
        if cols[i].button(label, key=f"nav_{k}", use_container_width=True):
            st.session_state.current_page = k
            st.rerun()
    if cols[-1].button("LOGOUT", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()
    st.divider()

# --- MODULES ---

if st.session_state.current_page == "Home Dashboard":
    page_header(); render_navigation_bar()
    st.title(f"Welcome, {st.session_state.username.title()}")
    f_df = get_data("families")
    st.metric("Total Devotees Enrolled", len(f_df))

elif st.session_state.current_page == "Enroll":
    page_header(); render_navigation_bar()
    st.header("Devotee Enrollment")
    # ... (Standard Enrollment Code) ...
    st.info("Module ready for devotee registration.")

elif st.session_state.current_page == "Billing":
    page_header(); render_navigation_bar(); st.header("Billing Desk")
    mode = st.radio("Mode", ["Enrolled Devotee", "Guest Devotee"], horizontal=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        s_name, s_id, s_addr, s_wa = "", 0, "", ""
        if mode == "Enrolled Devotee":
            f_data = get_data("families")
            if not f_data.empty:
                sel_f = st.selectbox("Select Family Head", f_data['head_name'].tolist())
                row = f_data[f_data['head_name'] == sel_f].iloc[0]
                s_name, s_id, s_addr, s_wa = row['head_name'], row['id'], row['address'], row['whatsapp']
        else:
            s_name = st.text_input("Guest Name *")
            s_addr = st.text_area("Address")
            s_wa = st.text_input("WhatsApp No")

        st.divider()
        servs = get_data("services")
        if not servs.empty:
            s_dict = {r['service_name']: r for _, r in servs.iterrows()}
            sel_s = st.selectbox("Select Service / Pooja", list(s_dict.keys()))
            srv = s_dict[sel_s]
            
            # --- DISPLAY BILL VALUE BEFORE GENERATING ---
            price = float(srv['price'])
            st.markdown(f"""
                <div style="background-color: #800000; padding: 20px; border-radius: 8px; border: 3px solid #FFD700; text-align: center; margin-bottom: 20px;">
                    <p style="color: #FFD700; font-size: 20px; font-weight: bold; margin: 0;">TOTAL PAYABLE AMOUNT</p>
                    <h1 style="color: #FFFFFF; font-size: 55px; margin: 10px 0;">₹ {price:,.2f}</h1>
                </div>
            """, unsafe_allow_html=True)
            
            m_no = st.text_input("Manual Bill No")
            b_no = st.text_input("Bill Book No")
            
            if st.button("Generate Receipt", use_container_width=True):
                if s_name:
                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    pld = {"family_id": s_id, "service_id": srv['id'], "amount": price, "date": now, "manual_bill_no": m_no, "bill_book_no": b_no, "guest_name": s_name if s_id == 0 else ""}
                    res = run_supabase_insert("transactions", pld)
                    if res and res.data:
                        pdf = generate_pdf(res.data[0]['id'], s_name, s_addr, sel_s, price, now, m_no, b_no)
                        st.success(f"Receipt Generated for ₹{price}!")
                        st.download_button("📥 Download PDF", pdf, f"Receipt_{res.data[0]['id']}.pdf")
                else: st.error("Please provide a name.")

elif st.session_state.current_page == "Samayavakuppu":
    page_header(); render_navigation_bar(); st.header("Samayavakuppu Student Bond Management")
    t1, t2 = st.tabs(["📝 Student Entry", "📋 View Records"])
    
    with t1:
        with st.form("student_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                nm = st.text_input("Student Name *")
                db = st.date_input("Date of Birth", value=None, min_value=MIN_DATE)
                fn = st.text_input("Father's Name")
                mob = st.text_input("Mobile No")
                adr = st.text_area("Address")
            with c2:
                bnk = st.text_input("Bond Issuing Bank")
                bno = st.text_input("Bond No")
                bis = st.date_input("Bond Issued Date", value=None)
                bex = st.date_input("Bond Expired Date", value=None)
                pht = st.file_uploader("Upload Student Photo", type=['jpg', 'png'])
            
            if st.form_submit_button("Save Record"):
                if nm and bno:
                    data = {"student_name": nm, "dob": format_date_for_db(db), "father_name": fn, "mobile_no": mob, "address": adr, "bond_bank": bnk, "bond_no": bno, "bond_issued_date": format_date_for_db(bis), "bond_expiry": format_date_for_db(bex), "photo": image_to_base64(pht)}
                    if run_supabase_insert("student_bonds", data): st.success("Student Bond Registered!"); st.rerun()
                else: st.error("Name and Bond No are required.")

    with t2:
        df = get_data("student_bonds")
        if not df.empty:
            sq = st.text_input("Search Name / Bond No")
            if sq: df = df[df['student_name'].str.contains(sq, case=False) | df['bond_no'].str.contains(sq)]
            
            opts = {f"{r['student_name']} (Bond: {r['bond_no']})": r['id'] for _, r in df.iterrows()}
            sel_s = st.selectbox("Select Student to View", [""] + list(opts.keys()))
            if sel_s:
                row = df[df['id'] == opts[sel_s]].iloc[0]
                with st.container(border=True):
                    v1, v2 = st.columns([1, 3])
                    with v1:
                        p_data = base64_to_image(row['photo'])
                        if p_data: st.image(p_data, use_container_width=True)
                        else: st.markdown("<div style='height:150px; background:#ddd; display:flex; align-items:center; justify-content:center;'>👤 No Photo</div>", unsafe_allow_html=True)
                    with v2:
                        st.subheader(f"Name: {row['student_name']}")
                        st.write(f"**Father:** {row['father_name']} | **DOB:** {format_date_for_ui(row['dob'])}")
                        st.write(f"**Contact:** {row['mobile_no']} | **Address:** {row['address']}")
                        st.divider()
                        st.markdown("#### Bond Information")
                        bc1, bc2, bc3 = st.columns(3)
                        bc1.info(f"**Bank:** {row['bond_bank']}\n\n**Bond No:** {row['bond_no']}")
                        bc2.write(f"**Issued On:**\n\n{format_date_for_ui(row.get('bond_issued_date'))}")
                        
                        # Reminder Logic
                        exp = safe_date_convert(row['bond_expiry'])
                        if exp and exp <= date.today() + timedelta(days=30):
                            bc3.error(f"**EXPIRY DATE:**\n\n{format_date_for_ui(row['bond_expiry'])}\n\n⚠️ URGENT: Expiring Soon")
                        else:
                            bc3.success(f"**Expiry Date:**\n\n{format_date_for_ui(row['bond_expiry'])}")

                if st.session_state.role == ADMIN_ROLE and st.button("🗑️ Delete Record"):
                    run_supabase_delete("student_bonds", row['id']); st.rerun()
        else: st.info("No records found.")

elif st.session_state.current_page == "Users":
    page_header(); render_navigation_bar(); st.header("User Management")
    if st.session_state.role == ADMIN_ROLE:
        u_df = get_data("users")
        st.dataframe(u_df[['username', 'role', 'rights']], use_container_width=True)
        # ... (User creation logic) ...

render_footer()
