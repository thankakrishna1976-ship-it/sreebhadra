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

def safe_date_convert(val):
    if not val or str(val).lower() in ["none", "nat", ""]: return None
    try: return pd.to_datetime(val).date()
    except: return None

def get_base64_of_bin_file(bin_file):
    try:
        if os.path.exists(bin_file):
            with open(bin_file, 'rb') as f: return base64.b64encode(f.read()).decode()
    except: pass
    return None

def image_to_base64(image_file):
    if image_file: return base64.b64encode(image_file.getvalue()).decode()
    return ""

def base64_to_image(base64_str):
    if base64_str:
        try: return io.BytesIO(base64.b64decode(base64_str))
        except: return None
    return None

def hash_password(password): return hashlib.sha256(password.encode()).hexdigest()

def verify_user(username, password):
    res = supabase.table('users').select('*').eq('username', username).execute()
    if res.data:
        user_data = res.data[0]
        if user_data['password_hash'] == hash_password(password):
            return True, user_data['role'], user_data.get('rights', 'Home Dashboard').split(',')
    return False, None, None

def to_excel(df):
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Report')
        return output.getvalue()
    except: return None

# --- PDF GENERATOR ---

def generate_pdf(receipt_no, devotee_name, devotee_address, service, amount, trans_date, manual_no, book_no):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    y = 780 
    try:
        if os.path.exists(LOGO_PATH):
            c.drawImage(ImageReader(LOGO_PATH), 50, y-50, width=50, height=50)
    except: pass
    c.setFont("Helvetica-Bold", 16); c.drawCentredString(300, y, TEMPLE_NAME_FULL)
    c.setFont("Helvetica", 10); c.drawCentredString(300, y-15, TRUST_DETAILS)
    c.drawCentredString(300, y-30, ADDRESS_LINE_1); c.drawCentredString(300, y-45, ADDRESS_LINE_2)
    c.line(50, y-60, 550, y-60)
    c.setFont("Helvetica-Bold", 12); c.drawString(50, y-90, f"RECEIPT No: #{receipt_no}"); c.drawString(400, y-90, f"DATE: {trans_date}")
    c.setFont("Helvetica", 11); c.drawString(50, y-130, f"Devotee: {devotee_name}"); c.drawString(50, y-150, f"Address: {devotee_address[:70]}")
    c.drawString(50, y-190, f"Seva: {service}"); c.setFont("Helvetica-Bold", 14); c.drawString(50, y-220, f"AMOUNT: Rs. {float(amount):,.2f}/-")
    c.save(); buffer.seek(0); return buffer

def generate_financial_pdf(income_df, expense_df, title, t_inc, t_exp, t_net):
    buffer = io.BytesIO(); doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet(); story = [Paragraph(TEMPLE_NAME_FULL, styles['Title']), Paragraph(title, styles['h3']), Spacer(1, 12)]
    data = [["Total Income", f"₹ {t_inc:,.2f}"], ["Total Expenses", f"₹ {t_exp:,.2f}"], ["Net Profit", f"₹ {t_net:,.2f}"]]
    t = Table(data, colWidths=[150, 150]); t.setStyle(TableStyle([('BACKGROUND', (0,0), (0,-1), colors.lightgrey), ('GRID', (0,0), (-1,-1), 0.5, colors.black)]))
    story.append(t); doc.build(story); buffer.seek(0); return buffer

# --- VISUAL COMPONENTS ---

def page_header():
    st.markdown("<style>.stToolbar {visibility: hidden;}</style>", unsafe_allow_html=True)
    c1, c2 = st.columns([1, 8])
    with c1: 
        try: st.image(LOGO_PATH, width=80)
        except: st.title("🕉️")
    with c2: st.markdown(f"<h1 style='color: #800000; border-bottom: 2px solid #b38728;'>{TEMPLE_NAME_FULL}</h1>", unsafe_allow_html=True)
    st.write("---")

def render_navigation_bar():
    ALL_PAGES = {"Home Dashboard": "HOME", "Enroll": "ENROLLMENT", "Search": "SEARCH", "Billing": "BILLING", "Expenses": "EXPENSES", "Reports": "REPORTS", "Assets": "ASSETS", "Samayavakuppu": "SAMAYAVAKUPPU", "Settings": "SETTINGS"}
    if st.session_state.role == ADMIN_ROLE:
        nav_items = ALL_PAGES.copy(); nav_items["Users"] = "USERS"
    else:
        nav_items = {k: v for k, v in ALL_PAGES.items() if k in st.session_state.get('rights', ["Home Dashboard"])}
    
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
        for _, r in df[df['dob'].astype(str).str.contains(today_md, na=False)].iterrows(): ticker.append(f"🎂 Happy Birthday: {r['head_name']}!")
        for _, r in df[df['yearly_pooja_date'].astype(str).str.contains(today_md, na=False)].iterrows(): ticker.append(f"🙏 Pooja Reminder: {r['head_name']}!")
    text = " | ".join(ticker) if ticker else "✨ Welcome to Sree Bhadreshwari Amman Temple Management System. ✨"
    st.markdown(f"""<style>.ticker-wrap {{ background: #800000; padding: 10px; border: 2px solid #FFD700; overflow: hidden; } .ticker {{ white-space: nowrap; animation: marquee 30s linear infinite; color: #FFD700; font-weight: bold; }} @keyframes marquee {{ 0% {{ transform: translateX(100%); }} 100% {{ transform: translateX(-100%); }} }}</style><div class="ticker-wrap"><div class="ticker">{text}</div></div><br>""", unsafe_allow_html=True)

# --- LOGIN & INIT ---

if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "Home Dashboard"
if 'new_family_id' not in st.session_state: st.session_state.new_family_id = None

if not st.session_state.logged_in:
    bg_64 = get_base64_of_bin_file(BACKGROUND_PATH)
    st.markdown(f"""<style>.stApp {{ background-image: url('data:image/jpg;base64,{bg_64}'); background-size: cover; background-position: center; }} label {{ color: #800000 !important; font-weight: bold; }}</style>""", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown("<div style='height: 15vh;'></div>", unsafe_allow_html=True)
        st.markdown("<h1 style='text-align: center; color: #800000;'>🕉️ Staff Login</h1>", unsafe_allow_html=True)
        un = st.text_input("Username")
        pw = st.text_input("Password", type="password")
        if st.button("SIGN IN", use_container_width=True):
            succ, role, rights = verify_user(un, pw)
            if succ:
                st.session_state.update({"logged_in": True, "username": un, "role": role, "rights": rights})
                st.rerun()
            else: st.error("Invalid credentials.")
    st.stop()

# --- MAIN APP UI ---

st.markdown("""<style>.stApp { background: linear-gradient(135deg, #bf953f 0%, #fcf6ba 50%, #aa771c 100%); }</style>""", unsafe_allow_html=True)
page_header()
render_navigation_bar()
render_news_ticker()

# --- MODULE ROUTING ---

if st.session_state.current_page == "Home Dashboard":
    st.title(f"Welcome, {st.session_state.username.title()}")
    df_t = get_data("transactions"); df_e = get_data("users_expenses")
    # Quick Dashboard Metrics Logic
    st.info("System Ready. Please use the navigation menu to manage devotees, billing, and reports.")

elif st.session_state.current_page == "Enroll":
    st.header("Devotee Enrollment")
    tab1, tab2 = st.tabs(["📝 Manual Entry", "📥 Bulk Upload"])
    with tab1:
        if st.session_state.new_family_id is None:
            with st.form("head_form"):
                n = st.text_input("Family Head Name *"); p = st.text_input("Phone *"); a = st.text_area("Address")
                d = st.date_input("DOB", value=None, min_value=MIN_DATE); w = st.date_input("Anniversary", value=None, min_value=MIN_DATE)
                s = st.selectbox("Star", [""] + NATCHATHIRAM_OPTIONS); pj = st.date_input("Yearly Pooja", value=None)
                if st.form_submit_button("Save Head"):
                    res = run_supabase_insert("families", {"head_name": n, "phone": p, "address": a, "dob": format_date_for_db(d), "wedding_date": format_date_for_db(w), "natchathiram": s, "yearly_pooja_date": format_date_for_db(pj)})
                    if res: st.session_state.new_family_id = res.data[0]['id']; st.rerun()
        else:
            st.success(f"Head Saved (ID: {st.session_state.new_family_id}). Add members or click reset.")
            if st.button("New Enrollment"): st.session_state.new_family_id = None; st.rerun()

elif st.session_state.current_page == "Billing":
    st.header("Billing Desk")
    mode = st.radio("Mode", ["Enrolled Devotee", "Guest Devotee"], horizontal=True)
    if mode == "Enrolled Devotee":
        fams = get_data("families"); mems = get_data("members")
        if not fams.empty:
            opts = {}
            for _, f in fams.iterrows():
                lbl = f"{f['head_name']} (Head) | 📱 {f['phone']}"
                opts[lbl] = {"id": f['id'], "name": f['head_name'], "addr": f['address'], "wa": f['whatsapp'] or f['phone']}
            for _, m in mems.iterrows():
                h_info = fams[fams['id'] == m['family_id']]
                hp = h_info['phone'].values[0] if not h_info.empty else ""
                ha = h_info['address'].values[0] if not h_info.empty else ""
                mp = m['phone'] if m['phone'] else hp
                lbl = f"{m['member_name']} ({m['relationship']}) | 📱 {mp}"
                opts[lbl] = {"id": m['family_id'], "name": m['member_name'], "addr": ha, "wa": m['whatsapp'] or mp}
            
            sel = st.selectbox("Search by Name or Mobile No", [""] + list(opts.keys()))
            if sel:
                d = opts[sel]; st.write(f"**Billing for:** {d['name']}")
                svs = get_data("services")
                if not svs.empty:
                    s_dict = {r['service_name']: r for _, r in svs.iterrows()}
                    s_sel = st.selectbox("Seva", list(s_dict.keys()))
                    if st.button("Generate Receipt"):
                        srv = s_dict[s_sel]
                        res = run_supabase_insert("transactions", {"family_id": d['id'], "service_id": srv['id'], "amount": srv['price'], "date": str(datetime.now())})
                        if res: st.success("Bill Saved!"); st.download_button("📥 PDF", generate_pdf(res.data[0]['id'], d['name'], d['addr'], s_sel, srv['price'], str(date.today()), "", ""), f"Rec_{res.data[0]['id']}.pdf")

elif st.session_state.current_page == "Reports":
    st.header("Financial Reports")
    report_mode = st.radio("Period:", ["Daily", "Weekly", "Monthly", "Custom"], horizontal=True)
    today = date.today(); start_d, end_d = today, today
    if report_mode == "Daily": start_d = st.date_input("Date", value=today); end_d = start_d
    elif report_mode == "Custom": c1, c2 = st.columns(2); start_d = c1.date_input("Start"); end_d = c2.date_input("End")
    
    df_t = get_data("transactions"); df_e = get_data("users_expenses")
    if not df_t.empty:
        df_t['dt'] = pd.to_datetime(df_t['date']).dt.date
        df_t = df_t[(df_t['dt'] >= start_d) & (df_t['dt'] <= end_d)]
    if not df_e.empty:
        df_e['dt'] = pd.to_datetime(df_e['payment_date']).dt.date
        df_e = df_e[(df_e['dt'] >= start_d) & (df_e['dt'] <= end_d)]

    ledger = []
    if not df_t.empty:
        for _, r in df_t.iterrows(): ledger.append({"Date": r['dt'], "Description": r['guest_name'] or "Income", "Income": r['amount'], "Expenses": 0, "Type": "Income"})
    if not df_e.empty:
        for _, r in df_e.iterrows(): ledger.append({"Date": r['dt'], "Description": r['expense_name'], "Income": 0, "Expenses": r['amount'], "Type": r['expense_type']})
    
    if ledger:
        df_l = pd.DataFrame(ledger).sort_values("Date")
        st.subheader("📊 Category-Wise Summary")
        cat_df = df_l.groupby('Type').agg({'Income': 'sum', 'Expenses': 'sum'}).reset_index()
        st.table(cat_df)
        
        st.subheader("📈 Distribution Chart")
        st.bar_chart(cat_df.set_index('Type')[['Income', 'Expenses']])
        
        st.subheader("📝 Detailed Ledger")
        st.dataframe(df_l, use_container_width=True)
    else: st.info("No data for this period.")

elif st.session_state.current_page == "Search":
    st.header("Search Devotees")
    df = get_data("families")
    if not df.empty:
        q = st.text_input("Search by Name or Phone")
        res = df[df['head_name'].str.contains(q, case=False) | df['phone'].str.contains(q)] if q else df
        st.dataframe(res, use_container_width=True)

elif st.session_state.current_page == "Samayavakuppu":
    st.header("Samayavakuppu Student Bond Management")
    t1, t2 = st.tabs(["📝 Entry", "📋 Records"])
    with t1:
        with st.form("bond_f"):
            sn = st.text_input("Student Name"); bn = st.text_input("Bond No")
            if st.form_submit_button("Save"):
                run_supabase_insert("student_bonds", {"student_name": sn, "bond_no": bn})
                st.success("Bond Registered.")

elif st.session_state.current_page == "Assets":
    st.header("Temple Assets")
    with st.form("asset_f"):
        an = st.text_input("Asset Name"); av = st.number_input("Value")
        if st.form_submit_button("Save Asset"):
            run_supabase_insert("assets", {"asset_name": an, "value": av}); st.rerun()
    st.dataframe(get_data("assets"), use_container_width=True)

elif st.session_state.current_page == "Settings":
    st.header("System Settings")
    t1, t2 = st.tabs(["Services", "Expenses"])
    with t1:
        with st.form("svc_f"):
            sn = st.text_input("Service Name"); sp = st.number_input("Price")
            if st.form_submit_button("Add"): run_supabase_insert("services", {"service_name": sn, "price": sp}); st.rerun()
        st.table(get_data("services"))

elif st.session_state.current_page == "Users":
    st.header("User Management")
    if st.session_state.role == ADMIN_ROLE:
        with st.form("user_f"):
            un = st.text_input("New Username"); up = st.text_input("Password", type="password")
            if st.form_submit_button("Create User"):
                run_supabase_insert("users", {"username": un, "password_hash": hash_password(up), "role": "user", "rights": "Home Dashboard"})
                st.rerun()
        st.dataframe(get_data("users", "id, username, role"), use_container_width=True)

st.markdown("""<style>.footer { position: fixed; left: 0; bottom: 0; width: 100%; background-color: #800000; color: #FFD700; text-align: center; padding: 10px 0; font-weight: bold; border-top: 2px solid #FFD700; }</style><div class="footer">Developed By : Sai Dharshini Info Solution</div>""", unsafe_allow_html=True)
