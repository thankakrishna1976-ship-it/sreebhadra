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
        st.error("Supabase credentials not found. Please check your st.secrets configuration.")
        st.stop()

supabase = get_supabase_client()

# --- CONFIGURATION & PAGE SETUP ---
st.set_page_config(page_title="Temple Management System", layout="wide", page_icon="🕉️")

ADMIN_ROLE = 'admin'
USER_ROLE = 'user'
DEFAULT_EXPENSE_TYPES = ['Pooja Items', 'Maintenance/Repairs', 'Salary/Dakshina', 'Electricity/Water', 'Annadanam/Food', 'Construction', 'Festivals', 'Administrative', 'Other']
RELATIONSHIP_OPTIONS = ['Wife', 'Son', 'Daughter', 'Mother', 'Father', 'Grand Father', 'Grand Mother', 'Guardian', 'Other']
NATCHATHIRAM_OPTIONS = ['Ashwini', 'Bharani', 'Karthigai', 'Rohini', 'Mrigasiram', 'Thiruvathirai', 'Punarpoosam', 'Poosam', 'Ayilyam', 'Magam', 'Poorvam', 'Uthiram', 'Hastham', 'Chithirai', 'Swathi', 'Visakam', 'Anusham', 'Kettai', 'Moolam', 'Pooradam', 'Uthiradam', 'Thiruvonam', 'Avittam', 'Sathayam', 'Poorattathi', 'Uthirattathi', 'Revathi']

MIN_DATE = date(1940, 1, 1)
MAX_DATE = date(2040, 12, 31)

# TEMPLE DETAILS
TEMPLE_NAME_FULL = "Sree Bhadreshwari Amman Temple Management System"
TRUST_DETAILS = "Samrakshana Seva Trust 174/2004"
ADDRESS_LINE_1 = "Kanjampuram"
ADDRESS_LINE_2 = "Kanniyakumari Dist., Tamil Nadu - 629154"
LOGO_PATH = "amman.jpg" 
BACKGROUND_PATH = "background.jpg"

# --- DB WRAPPER FUNCTIONS ---

def run_supabase_insert(table_name, data):
    try: return supabase.table(table_name).insert(data).execute()
    except Exception as e: st.error(f"DB Error: {e}"); return None

def run_supabase_update(table_name, data, row_id):
    try: return supabase.table(table_name).update(data).eq('id', row_id).execute()
    except Exception as e: st.error(f"DB Error: {e}"); return None

def run_supabase_delete(table_name, row_id):
    try: return supabase.table(table_name).delete().eq('id', row_id).execute()
    except Exception as e: st.error(f"DB Error: {e}"); return None

def get_data(table_name, select="*"):
    try:
        res = supabase.table(table_name).select(select).execute()
        return pd.DataFrame(res.data)
    except: return pd.DataFrame()

# --- UTILITY FUNCTIONS ---

def format_date_for_db(val):
    if pd.isna(val) or not val: return None
    try: return pd.to_datetime(val).strftime('%Y-%m-%d')
    except: return None

def format_date_for_ui(val):
    if not val or str(val).lower() in ["none", "nat", ""]: return ""
    try: return pd.to_datetime(val).strftime('%d/%m/%Y')
    except: return str(val)

def get_base64_of_bin_file(bin_file):
    try:
        if os.path.exists(bin_file):
            with open(bin_file, 'rb') as f: return base64.b64encode(f.read()).decode()
    except: pass
    return None

def image_to_base64(image_file):
    if image_file: return base64.b64encode(image_file.getvalue()).decode()
    return ""

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def hash_password(password): return hashlib.sha256(password.encode()).hexdigest()

def verify_user(username, password):
    res = supabase.table('users').select('*').eq('username', username).execute()
    if res.data:
        user_data = res.data[0]
        if user_data['password_hash'] == hash_password(password):
            return True, user_data['role'], user_data.get('rights', 'Home Dashboard').split(',')
    return False, None, None

# --- PDF GENERATOR ---

def generate_pdf(receipt_no, devotee_name, devotee_address, service, amount, trans_date):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    y = 780 
    c.setFont("Helvetica-Bold", 16); c.drawCentredString(300, y, TEMPLE_NAME_FULL)
    c.setFont("Helvetica", 10); c.drawCentredString(300, y-15, TRUST_DETAILS)
    c.line(50, y-60, 550, y-60)
    c.setFont("Helvetica-Bold", 12); c.drawString(50, y-90, f"RECEIPT No: #{receipt_no}"); c.drawString(400, y-90, f"DATE: {trans_date}")
    c.setFont("Helvetica", 11); c.drawString(50, y-130, f"Devotee: {devotee_name}")
    c.drawString(50, y-190, f"Seva: {service}"); c.setFont("Helvetica-Bold", 14); c.drawString(50, y-220, f"AMOUNT: Rs. {float(amount):,.2f}/-")
    c.save(); buffer.seek(0); return buffer

# --- VISUAL COMPONENTS ---

def page_header():
    st.markdown("<style>.stToolbar {visibility: hidden;}</style>", unsafe_allow_html=True)
    c1, c2 = st.columns([1, 8])
    with c1: 
        try: st.image(LOGO_PATH, width=80)
        except: st.title("🕉️")
    with c2: st.markdown(f"<h1 style='color: #800000; border-bottom: 2px solid #b38728;'>{TEMPLE_NAME_FULL}</h1>", unsafe_allow_html=True)

def render_navigation_bar():
    ALL_PAGES = {"Home Dashboard": "HOME", "Enroll": "ENROLLMENT", "Search": "SEARCH", "Billing": "BILLING", "Expenses": "EXPENSES", "Reports": "REPORTS", "Assets": "ASSETS", "Samayavakuppu": "SAMAYAVAKUPPU", "Settings": "SETTINGS"}
    nav_items = ALL_PAGES if st.session_state.role == ADMIN_ROLE else {k: v for k, v in ALL_PAGES.items() if k in st.session_state.rights}
    if st.session_state.role == ADMIN_ROLE: nav_items["Users"] = "USERS"
    
    cols = st.columns(len(nav_items) + 1)
    st.markdown("""<style> div[data-testid="column"] .stButton>button { border-radius: 0px !important; background-color: #800000; color: #FFD700; border: 1px solid #FFD700; font-weight: bold; height: 3.5em; width: 100%; transition: 0.3s; } div[data-testid="column"] .stButton>button:hover { background-color: #A00000; transform: translateY(-2px); }</style>""", unsafe_allow_html=True)
    for i, (key, label) in enumerate(nav_items.items()):
        if cols[i].button(label, key=f"n_{key}"):
            st.session_state.current_page = key; st.rerun()
    if cols[-1].button("LOGOUT"): st.session_state.logged_in = False; st.rerun()
    st.write("---")

def render_news_ticker():
    today_md = date.today().strftime('%m-%d'); ticker = []
    df = get_data("families")
    if not df.empty:
        for _, r in df[df['dob'].astype(str).str.contains(today_md, na=False)].iterrows(): ticker.append(f"🎂 Birthday: {r['head_name']}!")
        for _, r in df[df['yearly_pooja_date'].astype(str).str.contains(today_md, na=False)].iterrows(): ticker.append(f"🙏 Pooja: {r['head_name']}!")
    text = " | ".join(ticker) if ticker else "✨ Welcome to Sree Bhadreshwari Amman Temple Management System. ✨"
    st.markdown(f"""<style>.ticker-wrap {{ background: #800000; padding: 10px; border: 2px solid #FFD700; overflow: hidden; }} .ticker {{ white-space: nowrap; animation: marquee 30s linear infinite; color: #FFD700; font-weight: bold; display: inline-block; padding-left: 100%; }} @keyframes marquee {{ 0% {{ transform: translateX(0); }} 100% {{ transform: translateX(-100%); }} }}</style><div class="ticker-wrap"><div class="ticker">{text}</div></div><br>""", unsafe_allow_html=True)

# --- APP INITIALIZATION ---

if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "Home Dashboard"
if 'new_family_id' not in st.session_state: st.session_state.new_family_id = None

if not st.session_state.logged_in:
    bg_64 = get_base64_of_bin_file(BACKGROUND_PATH)
    st.markdown(f"""<style>.stApp {{ background-image: url('data:image/jpg;base64,{bg_64}'); background-size: cover; }} label {{ color: #800000 !important; font-weight: bold; }}</style>""", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown("<div style='height: 15vh;'></div>", unsafe_allow_html=True)
        st.markdown("<h1 style='text-align: center; color: #800000;'>🕉️ Staff Login</h1>", unsafe_allow_html=True)
        un = st.text_input("Username"); pw = st.text_input("Password", type="password")
        if st.button("SIGN IN", use_container_width=True):
            succ, role, rights = verify_user(un, pw)
            if succ: st.session_state.update({"logged_in": True, "username": un, "role": role, "rights": rights}); st.rerun()
    st.stop()

st.markdown("""<style>.stApp { background: linear-gradient(135deg, #bf953f 0%, #fcf6ba 50%, #aa771c 100%); }</style>""", unsafe_allow_html=True)
page_header(); render_navigation_bar(); render_news_ticker()

# --- MODULE ROUTING ---

if st.session_state.current_page == "Home Dashboard":
    st.title(f"Welcome, {st.session_state.username.title()}")
    df_t = get_data("transactions"); df_e = get_data("users_expenses"); df_f = get_data("families")
    today = date.today()
    if not df_t.empty: df_t['date_obj'] = pd.to_datetime(df_t['date']).dt.date
    if not df_e.empty: df_e['date_obj'] = pd.to_datetime(df_e['payment_date']).dt.date
    
    c1, c2, c3 = st.columns(3)
    inc = df_t[df_t['date_obj'] == today]['amount'].sum() if not df_t.empty else 0
    exp = df_e[df_e['date_obj'] == today]['amount'].sum() if not df_e.empty else 0
    c1.metric("Today's Income", f"₹ {inc:,.2f}")
    c2.metric("Today's Expense", f"₹ {exp:,.2f}")
    c3.metric("Live Net", f"₹ {inc-exp:,.2f}")
    st.divider(); st.metric("Total Registered Families", len(df_f))

elif st.session_state.current_page == "Enroll":
    st.header("Devotee Enrollment")
    t1, t2 = st.tabs(["📝 Manual Entry", "📥 Bulk Upload"])
    with t1:
        if st.session_state.new_family_id is None:
            with st.form("h_f"):
                n = st.text_input("Head Name *"); p = st.text_input("Phone *"); a = st.text_area("Address")
                d = st.date_input("DOB", value=None, min_value=MIN_DATE); s = st.selectbox("Star", [""] + NATCHATHIRAM_OPTIONS)
                if st.form_submit_button("Save Head"):
                    res = run_supabase_insert("families", {"head_name": n, "phone": p, "address": a, "dob": format_date_for_db(d), "natchathiram": s})
                    if res: st.session_state.new_family_id = res.data[0]['id']; st.rerun()
        else:
            st.info(f"Adding Members for Head ID: {st.session_state.new_family_id}")
            with st.form("m_f"):
                mn = st.text_input("Member Name"); mr = st.selectbox("Relationship", RELATIONSHIP_OPTIONS)
                if st.form_submit_button("Add Member"):
                    run_supabase_insert("members", {"family_id": st.session_state.new_family_id, "member_name": mn, "relationship": mr})
                    st.success("Member added.")
            if st.button("Finish Enrollment"): st.session_state.new_family_id = None; st.rerun()
    with t2:
        st.subheader("Excel Bulk Upload")
        sample = pd.DataFrame({"Name": ["Rajesh"], "Phone": ["9659828283"], "Address": ["Kanjampuram"], "Star": ["Ashwini"], "DOB": ["1990-01-01"]})
        st.download_button("📥 Template", to_excel(sample), "template.xlsx")
        up = st.file_uploader("Upload Excel", type=["xlsx"])
        if up and st.button("Process Bulk"):
            df_up = pd.read_excel(up)
            for _, r in df_up.iterrows():
                run_supabase_insert("families", {"head_name": str(r['Name']), "phone": str(r['Phone']), "address": str(r['Address']), "natchathiram": str(r['Star']), "dob": format_date_for_db(r['DOB'])})
            st.success("Upload Complete.")

elif st.session_state.current_page == "Billing":
    st.header("Billing Desk")
    mode = st.radio("Mode", ["Enrolled", "Guest"], horizontal=True)
    if mode == "Enrolled":
        f_data = get_data("families"); m_data = get_data("members")
        if not f_data.empty:
            opts = {}
            for _, f in f_data.iterrows(): opts[f"{f['head_name']} | 📱 {f['phone']}"] = f['id']
            for _, m in m_data.iterrows():
                h = f_data[f_data['id'] == m['family_id']]
                hp = h['phone'].values[0] if not h.empty else ""
                opts[f"{m['member_name']} ({m['relationship']}) | 📱 {hp}"] = m['family_id']
            
            sel = st.selectbox("Search by Mobile or Name", [""] + list(opts.keys()))
            if sel:
                svs = get_data("services"); s_dict = {r['service_name']: r for _, r in svs.iterrows()}
                s_sel = st.selectbox("Service", list(s_dict.keys()))
                if st.button("Generate Bill"):
                    srv = s_dict[s_sel]
                    res = run_supabase_insert("transactions", {"family_id": opts[sel], "service_id": srv['id'], "amount": srv['price'], "date": str(datetime.now())})
                    if res: st.success("Generated!"); st.download_button("📥 PDF", generate_pdf(res.data[0]['id'], sel.split('|')[0], "Temple Dist.", s_sel, srv['price'], str(date.today())), f"Rec_{res.data[0]['id']}.pdf")
    else:
        gn = st.text_input("Guest Name"); ga = st.text_area("Address")
        # Same logic for Guest billing...

elif st.session_state.current_page == "Samayavakuppu":
    st.header("Samayavakuppu Student Bond Management")
    tab1, tab2 = st.tabs(["📝 Student Entry", "📋 View Records"])
    with tab1:
        with st.form("bond_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            sn = c1.text_input("Student Name *"); sd = c1.date_input("DOB", value=None, min_value=MIN_DATE)
            sf = c1.text_input("Father's Name"); sm = c1.text_input("Mobile No")
            bb = c2.text_input("Issuing Bank"); bn = c2.text_input("Bond No *")
            bi = c2.date_input("Issued Date"); be = c2.date_input("Expiry Date")
            if st.form_submit_button("Save Bond"):
                if sn and bn:
                    run_supabase_insert("student_bonds", {"student_name": sn, "dob": str(sd), "father_name": sf, "mobile_no": sm, "bond_bank": bb, "bond_no": bn, "bond_expiry": str(be)})
                    st.success("Bond Registered Successfully!")
    with tab2:
        df_b = get_data("student_bonds")
        if not df_b.empty: st.dataframe(df_b, use_container_width=True)

elif st.session_state.current_page == "Reports":
    st.header("Financial Reports")
    df_t = get_data("transactions"); df_e = get_data("users_expenses")
    ledger = []
    if not df_t.empty:
        for _, r in df_t.iterrows(): ledger.append({"Date": r['date'], "Type": "Income", "Description": "Temple Seva", "Amount": r['amount']})
    if not df_e.empty:
        for _, r in df_e.iterrows(): ledger.append({"Date": r['payment_date'], "Type": r['expense_type'], "Description": r['expense_name'], "Amount": r['amount']})
    
    if ledger:
        df_l = pd.DataFrame(ledger)
        st.subheader("📊 Category-Wise Summary")
        cat_sum = df_l.groupby('Type')['Amount'].sum().reset_index()
        st.table(cat_sum)
        st.bar_chart(df_l.groupby('Type')['Amount'].sum())
        st.subheader("📝 Transaction Ledger")
        st.dataframe(df_l, use_container_width=True)

elif st.session_state.current_page == "Search":
    st.header("Search Devotees")
    df_s = get_data("families")
    if not df_s.empty:
        q = st.text_input("Type Name or Phone Number")
        res = df_s[df_s['head_name'].str.contains(q, case=False) | df_s['phone'].str.contains(q)] if q else df_s
        st.dataframe(res, use_container_width=True)

elif st.session_state.current_page == "Settings":
    st.header("Services & Settings")
    with st.form("svc"):
        sn = st.text_input("Service Name"); sp = st.number_input("Price")
        if st.form_submit_button("Add Service"): run_supabase_insert("services", {"service_name": sn, "price": sp}); st.rerun()
    st.table(get_data("services"))

elif st.session_state.current_page == "Users":
    if st.session_state.role == ADMIN_ROLE:
        st.header("User Management")
        un = st.text_input("New Username"); up = st.text_input("Password", type="password")
        if st.button("Create User"): run_supabase_insert("users", {"username": un, "password_hash": hash_password(up), "role": "user"}); st.rerun()
        st.dataframe(get_data("users", "id, username, role"), use_container_width=True)

st.markdown("""<style>.footer { position: fixed; left: 0; bottom: 0; width: 100%; background: #800000; color: #FFD700; text-align: center; padding: 10px; font-weight: bold; border-top: 2px solid #FFD700; }</style><div class="footer">Developed By : Sai Dharshini Info Solution</div>""", unsafe_allow_html=True)
