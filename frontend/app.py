import os, requests, re, tempfile, msoffcrypto, base64
from textwrap import dedent
from io import BytesIO
from pathlib import Path
from datetime import datetime, date
from dotenv import load_dotenv
import streamlit as st
from openpyxl import load_workbook
import plotly.express as px
import plotly.graph_objects as go

# Create assets folder if not exists
ASSETS_DIR = Path(__file__).resolve().parent / "assets"
ASSETS_DIR.mkdir(parents=True, exist_ok=True)

def get_asset_base64(filename):
    """Retrieve an asset file (logo, background) as Base64 data URL if present.
    Searches multiple candidate directories to handle different Streamlit launch paths.
    """
    stem = Path(filename).stem
    names_to_try = [filename, f"{stem}.png", f"{stem}.jpg", f"{stem}.jpeg", f"{stem}.svg"]
    
    candidates_dirs = [
        ASSETS_DIR,
        Path(__file__).resolve().parent / "assets",
        Path.cwd() / "assets",
        Path.cwd() / "frontend" / "assets",
    ]
    
    for name in names_to_try:
        for directory in candidates_dirs:
            file_path = directory / name
            if file_path.exists():
                try:
                    with open(file_path, "rb") as f:
                        data = f.read()
                    ext = file_path.suffix.lstrip(".").lower()
                    mime = (
                        "image/svg+xml" if ext == "svg"
                        else "image/png" if ext == "png"
                        else "image/jpeg"
                    )
                    return f"data:{mime};base64,{base64.b64encode(data).decode()}"
                except Exception:
                    pass
    return None


# Load environment variables
load_dotenv(Path(__file__).resolve().parents[1] / ".env")
load_dotenv(Path(__file__).resolve().parent / ".env")
load_dotenv()

st.set_page_config(
    page_title="Attendance Report PORTAL - Login",
    page_icon="⏱️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

API = os.getenv("API_URL", "http://localhost:8000").rstrip("/")

# Initialize session state variables
if "token" not in st.session_state:
    st.session_state.token = None
if "user" not in st.session_state:
    st.session_state.user = "coordinator"
if "role" not in st.session_state:
    st.session_state.role = "coordinator"
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "dashboard"
if "history" not in st.session_state:
    st.session_state.history = []
if "analytics_data" not in st.session_state:
    st.session_state.analytics_data = None
if "analytics_file_name" not in st.session_state:
    st.session_state.analytics_file_name = None
# Reports from DB
if "available_reports" not in st.session_state:
    st.session_state.available_reports = None
if "selected_report_uuid" not in st.session_state:
    st.session_state.selected_report_uuid = None




def clean_html(s: str) -> str:
    """Strip all line indentation to prevent Streamlit Markdown from treating HTML as code blocks."""
    return re.sub(r'^\s+', '', s, flags=re.MULTILINE).strip()


def inject_custom_css(login_page=False):
    """Inject application-wide CSS, hiding Streamlit top headers and styling Attendance login page."""
    if login_page:
        st.markdown(clean_html("""
        <style>
        /* ============================================================
           LOGIN PAGE FULL VIEWPORT (ZERO MARGINS, ZERO BLACK SPACE)
           ============================================================ */
        header,
        header[data-testid="stHeader"],
        [data-testid="stHeader"],
        .stAppHeader,
        .stAppDeployButton,
        div[data-testid="stToolbar"],
        #MainMenu,
        [data-testid="stDecoration"],
        footer {
            display: none !important;
            visibility: hidden !important;
            height: 0 !important;
            width: 0 !important;
            opacity: 0 !important;
            pointer-events: none !important;
        }

        * {
            box-sizing: border-box;
        }

        html, body {
            margin: 0 !important;
            padding: 0 !important;
            width: 100vw !important;
            height: 100vh !important;
            overflow: hidden !important;
            background: #ffffff !important;
        }

        .stApp,
        [data-testid="stAppViewContainer"],
        section.main,
        [data-testid="stMainBlockContainer"],
        .block-container {
            margin: 0 !important;
            padding: 0 !important;
            width: 100vw !important;
            max-width: 100vw !important;
            height: 100vh !important;
            max-height: 100vh !important;
            overflow: hidden !important;
            top: 0 !important;
            background: #ffffff !important;
        }

        div[data-testid="stVerticalBlock"] {
            gap: 0 !important;
            padding: 0 !important;
            margin: 0 !important;
        }

        div[data-testid="stHorizontalBlock"] {
            height: 100vh !important;
            min-height: 100vh !important;
            margin: 0 !important;
            padding: 0 !important;
            gap: 0 !important;
            width: 100vw !important;
            display: flex !important;
        }

        /* Remove header link icons / anchor links completely */
        a.header-anchor,
        [data-testid="stHeaderActionElements"],
        .stMarkdown a.header-anchor,
        h1 a, h2 a, h3 a {
            display: none !important;
            visibility: hidden !important;
            pointer-events: none !important;
            opacity: 0 !important;
            width: 0 !important;
            height: 0 !important;
        }

        /* Left Branding Column (75% Width, Full 100vh Height) */
        div[data-testid="column"]:nth-child(1),
        div[data-testid="stColumn"]:nth-child(1) {
            height: 100vh !important;
            padding: 0 !important;
            flex: 0 0 75% !important;
            width: 75% !important;
            max-width: 75% !important;
            margin: 0 !important;
            overflow: hidden !important;
        }

        /* Right Form Column (25% Width, Pure White Theme, Positioned Above Center) */
        div[data-testid="column"]:nth-child(2),
        div[data-testid="stColumn"]:nth-child(2) {
            height: 100vh !important;
            min-height: 100vh !important;
            background: #ffffff !important;
            background-color: #ffffff !important;
            display: flex !important;
            flex-direction: column !important;
            align-items: center !important;
            justify-content: flex-start !important;
            padding-top: 18vh !important;
            padding-left: 20px !important;
            padding-right: 20px !important;
            padding-bottom: 30px !important;
            box-shadow: -10px 0 30px rgba(0, 0, 0, 0.08) !important;
            z-index: 10 !important;
            flex: 0 0 25% !important;
            width: 25% !important;
            max-width: 25% !important;
            min-width: 280px !important;
            margin: 0 !important;
        }

        div[data-testid="column"]:nth-child(2) > div,
        div[data-testid="stColumn"]:nth-child(2) > div,
        div[data-testid="column"]:nth-child(2) [data-testid="stVerticalBlock"],
        div[data-testid="stColumn"]:nth-child(2) [data-testid="stVerticalBlock"] {
            width: 100% !important;
            max-width: 320px !important;
            display: flex !important;
            flex-direction: column !important;
            align-items: center !important;
            justify-content: flex-start !important;
            margin: 0 auto !important;
        }

        /* Branding Section */
        .branding-section {
            height: 100vh;
            width: 100%;
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            align-items: center;
            padding: 50px 40px 30px;
            color: white;
            position: relative;
            overflow: hidden;
        }

        .branding-section::before {
            content: "";
            position: absolute;
            inset: 0;
            background: rgba(0, 50, 75, 0.15);
            z-index: 0;
        }

        .branding-content {
            display: flex;
            flex-direction: column;
            align-items: center;
            text-align: center;
            flex: 1;
            justify-content: center;
            z-index: 2;
            width: 100%;
        }

        .logo-container {
            margin-bottom: 25px;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            width: 100% !important;
        }

        .logo,
        .branding-section img {
            display: block !important;
            width: 140px !important;
            height: 140px !important;
            max-width: 140px !important;
            max-height: 140px !important;
            object-fit: contain !important;
            margin: 0 auto !important;
            opacity: 1 !important;
            visibility: visible !important;
            filter: drop-shadow(0 4px 12px rgba(0, 0, 0, 0.25)) !important;
        }

        .institution-info {
            margin-bottom: 30px;
        }

        .institution-name {
            font-size: 26px;
            font-weight: 300;
            line-height: 1.3;
            margin: 0 0 6px;
            color: white;
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
        }

        .portal-title {
            margin-top: 15px;
        }

        .portal-title-text {
            margin: 0;
            font-size: 30px;
            font-weight: 700;
            letter-spacing: 2px;
            color: white;
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.25);
        }

        /* Footer */
        .branding-footer {
            width: 100%;
            text-align: center;
            z-index: 2;
        }

        .social-icons {
            display: flex;
            justify-content: center;
            gap: 18px;
            margin-bottom: 25px;
        }

        .social-icon {
            width: 38px;
            height: 38px;
            background: rgba(255, 255, 255, 0.12);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white !important;
            text-decoration: none !important;
            transition: all 0.3s ease;
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.25);
        }

        .social-icon:hover {
            background: rgba(255, 255, 255, 0.3);
            transform: translateY(-2px);
        }

        .social-icon svg {
            width: 18px;
            height: 18px;
            stroke: currentColor;
        }

        .copyright {
            font-size: 12px;
            line-height: 1.5;
            opacity: 0.9;
            color: white;
        }

        .copyright p {
            margin: 0;
        }

        .support-text {
            margin-top: 4px !important;
            opacity: 0.8;
        }

        /* ============================================================
           PURE WHITE THEME LOGIN FORM STYLES
           ============================================================ */

        [data-testid="stForm"] {
            border: none !important;
            padding: 0 !important;
            background: #ffffff !important;
            background-color: #ffffff !important;
            box-shadow: none !important;
            width: 100% !important;
            max-width: 340px !important;
            margin: 0 auto !important;
        }

        div[data-testid="stTextInput"],
        [data-testid="stForm"] [data-testid="element-container"] {
            margin-bottom: 22px !important;
            width: 100% !important;
        }

        div[data-baseweb="input"] {
            background: #f8fafc !important;
            background-color: #f8fafc !important;
            border: 1.5px solid #cbd5e1 !important;
            border-radius: 12px !important;
            min-height: 52px !important;
            box-shadow: none !important;
            transition: all 0.2s ease-in-out !important;
            overflow: hidden !important;
        }

        div[data-baseweb="input"] div,
        div[data-baseweb="base-input"],
        div[data-baseweb="base-input"] div {
            background: transparent !important;
            background-color: transparent !important;
            border: none !important;
            box-shadow: none !important;
        }

        div[data-baseweb="input"]:focus-within {
            border: 2px solid #3b82f6 !important;
            border-color: #3b82f6 !important;
            background: #f8fafc !important;
            background-color: #f8fafc !important;
            box-shadow: 0 0 0 4px rgba(59, 130, 246, 0.22) !important;
        }

        div[data-baseweb="input"] input,
        div[data-baseweb="base-input"] input,
        div[data-testid="stTextInput"] input {
            color: #1e293b !important;
            background: transparent !important;
            background-color: transparent !important;
            font-size: 16px !important;
            font-weight: 500 !important;
            padding: 14px 18px !important;
            font-family: 'Inter', system-ui, -apple-system, sans-serif !important;
        }

        div[data-baseweb="input"] input::placeholder,
        div[data-baseweb="base-input"] input::placeholder,
        div[data-testid="stTextInput"] input::placeholder {
            color: #64748b !important;
            opacity: 1 !important;
            font-weight: 400 !important;
        }

        div[data-baseweb="input"] button,
        div[data-baseweb="input"] svg {
            background: transparent !important;
            background-color: transparent !important;
            color: #64748b !important;
        }

        input:-webkit-autofill,
        input:-webkit-autofill:hover, 
        input:-webkit-autofill:focus, 
        input:-webkit-autofill:active {
            -webkit-box-shadow: 0 0 0 1000px #f8fafc inset !important;
            -webkit-text-fill-color: #1e293b !important;
            transition: background-color 5000s ease-in-out 0s !important;
        }

        div[data-testid="stFormSubmitButton"] button,
        button[data-testid="baseButton-primary"] {
            width: 100% !important;
            min-height: 52px !important;
            padding: 14px !important;
            background: #2b8cee !important;
            background-color: #2b8cee !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 10px !important;
            font-size: 17px !important;
            font-weight: 600 !important;
            cursor: pointer !important;
            transition: all 0.2s ease !important;
            margin-top: 10px !important;
            box-shadow: 0 4px 14px rgba(43, 140, 238, 0.35) !important;
        }

        div[data-testid="stFormSubmitButton"] button:hover,
        button[data-testid="baseButton-primary"]:hover {
            background: #1e7bdc !important;
            background-color: #1e7bdc !important;
            transform: translateY(-1px);
            box-shadow: 0 6px 18px rgba(43, 140, 238, 0.45) !important;
        }

        div[data-testid="column"]:nth-child(2) div[data-testid="stFormSubmitButton"] button:active {
            transform: translateY(0);
        }

        .login-links {
            margin-top: 20px;
            text-align: center;
        }

        .forgot-password {
            color: #0077ff ;
            text-decoration: none ;
            font-size: 14px;
            font-weight: 500;
            display: block;
            transition: color 0.2s ease;
        }

        .forgot-password:hover {
            color: #0056b3 ;
            text-decoration: underline ;
        }
        </style>
        """), unsafe_allow_html=True)
    else:
        st.markdown(clean_html("""
        <style>
        header,
        header[data-testid="stHeader"],
        [data-testid="stHeader"],
        .stAppHeader,
        .stAppDeployButton,
        div[data-testid="stToolbar"],
        #MainMenu,
        [data-testid="stDecoration"],
        footer {
            display: none ;
            visibility: hidden ;
            height: 0 ;
            width: 0 ;
        }

        * {
            box-sizing: border-box;
        }

        .stApp {
            background-color: #0b1329;
            color: #f1f5f9;
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
        }

        .block-container {
            padding: 1.5rem 2rem ;
            max-width: 100% ;
        }

        .nav-container {
            display: flex;
            align-items: center;
            justify-content: space-between;
            background: rgba(15, 23, 42, 0.9);
            backdrop-filter: blur(14px);
            border: 1px solid rgba(0, 180, 216, 0.2);
            padding: 0.75rem 1.5rem;
            border-radius: 14px;
            margin-bottom: 1.5rem;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.4);
        }

        .glass-card {
            background: rgba(15, 23, 42, 0.75);
            backdrop-filter: blur(12px);
            border: 1px solid rgba(0, 180, 216, 0.15);
            border-radius: 16px;
            padding: 1.5rem;
            margin-bottom: 1.25rem;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.35);
            transition: transform 0.2s ease, border-color 0.2s ease;
        }

        .glass-card:hover {
            border-color: rgba(0, 180, 216, 0.5);
        }

        .stat-box {
            background: linear-gradient(
                145deg,
                rgba(15, 23, 42, 0.95),
                rgba(11, 19, 41, 0.95)
            );
            border: 1px solid rgba(0, 180, 216, 0.25);
            border-radius: 14px;
            padding: 1.25rem;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.25);
        }

        .stat-val {
            font-size: 1.8rem;
            font-weight: 800;
            color: #38bdf8;
        }

        .stat-label {
            font-size: 0.85rem;
            color: #94a3b8;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-top: 0.25rem;
        }

        .stButton > button {
            border-radius: 10px ;
            font-weight: 600 ;
            transition: all 0.2s ease-in-out ;
        }

        .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #00b4d8 0%, #0077b6 100%) ;
            border: none ;
            color: #ffffff ;
            box-shadow: 0 4px 14px rgba(0, 119, 182, 0.4) ;
        }

        .stButton > button[kind="primary"]:hover {
            background: linear-gradient(135deg, #0096c7 0%, #005f8a 100%) ;
            transform: translateY(-1px);
            box-shadow: 0 6px 18px rgba(0, 119, 182, 0.5) ;
        }

        .stButton > button[kind="secondary"] {
            background: rgba(255, 255, 255, 0.06) ;
            border: 1px solid rgba(255, 255, 255, 0.15) ;
            color: #cbd5e1 ;
        }

        .stButton > button[kind="secondary"]:hover {
            border-color: rgba(0, 180, 216, 0.5) ;
            color: #38bdf8 ;
        }

        div[data-baseweb="input"] {
            background-color: rgba(15, 23, 42, 0.6) ;
            border: 1px solid rgba(255, 255, 255, 0.15) ;
            border-radius: 8px ;
            color: #ffffff ;
        }

        div[data-baseweb="input"]:focus-within {
            border-color: #0077b6 ;
            box-shadow: 0 0 0 3px rgba(0, 119, 182, 0.25) ;
        }

        .user-badge {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.4rem 0.85rem;
            background: rgba(0, 180, 216, 0.15);
            border: 1px solid rgba(0, 180, 216, 0.3);
            border-radius: 8px;
            font-size: 0.88rem;
            font-weight: 600;
            color: #38bdf8;
        }
        </style>
        """), unsafe_allow_html=True)


def check_backend_health():
    try:
        r = requests.get(f"{API}/health", timeout=3)
        return r.ok
    except Exception:
        return False

def parse_date_header(v):
    if v is None: return None
    if isinstance(v, (datetime, date)): return v.date()
    s = re.sub(r"\s+", " ", str(v).strip())
    for f in ["%d %b %y", "%d %b %Y", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%Y-%m-%d", "%Y/%m/%d", "%d %B %y", "%d %B %Y"]:
        try: return datetime.strptime(s, f).date()
        except ValueError: pass
    return None

def decrypt_excel_bytes(file_bytes, password):
    """Decrypt encrypted Excel bytes using disk-backed temporary files matching backend excel_service."""
    fd_in, in_path = tempfile.mkstemp(suffix=".xlsx")
    fd_out, out_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd_in)
    os.close(fd_out)

    try:
        with open(in_path, "wb") as f_in:
            f_in.write(file_bytes)

        with open(in_path, "rb") as f_in:
            office_file = msoffcrypto.OfficeFile(f_in)
            # Try raw password and trimmed password
            passwords_to_try = [password]
            if password.strip() != password:
                passwords_to_try.append(password.strip())

            last_err = None
            decrypted_success = False
            for pwd in passwords_to_try:
                try:
                    f_in.seek(0)
                    office_file.load_key(password=pwd)
                    with open(out_path, "wb") as f_out:
                        office_file.decrypt(f_out)
                    decrypted_success = True
                    break
                except Exception as e:
                    last_err = e

            if not decrypted_success:
                raise last_err or ValueError("Decryption failed.")

        with open(out_path, "rb") as f_out:
            return f_out.read()
    finally:
        if os.path.exists(in_path):
            try: os.remove(in_path)
            except Exception: pass
        if os.path.exists(out_path):
            try: os.remove(out_path)
            except Exception: pass

def load_excel_workbook(file_bytes, password=""):
    """Load an Excel workbook, performing msoffcrypto decryption if password-protected."""
    # 1. Try loading directly first (works for unencrypted .xlsx and .xlsm)
    for data_only in [False, True]:
        try:
            return load_workbook(filename=BytesIO(file_bytes), data_only=data_only)
        except Exception:
            pass

    # 2. Check if password was provided
    if not password:
        raise ValueError("🔒 This Excel file is password-protected. Please enter the password in the password field and click 'Run Analysis'.")

    # 3. Attempt disk-backed msoffcrypto decryption
    try:
        decrypted_bytes = decrypt_excel_bytes(file_bytes, password)
    except msoffcrypto.exceptions.InvalidKeyError:
        raise ValueError("❌ Password incorrect. The password entered does not unlock this Excel file.")
    except Exception as decrypt_err:
        raise ValueError(f"❌ Could not decrypt Excel file: {decrypt_err}") from decrypt_err

    # 4. Load decrypted workbook
    for data_only in [False, True]:
        try:
            return load_workbook(filename=BytesIO(decrypted_bytes), data_only=data_only)
        except Exception:
            pass

    raise ValueError("Decrypted Excel file could not be parsed by openpyxl.")



def parse_swipe_dt(v):
    if v is None: return None
    if isinstance(v, datetime): return v
    if isinstance(v, date): return datetime.combine(v, datetime.min.time())
    s = re.sub(r"\s+", " ", str(v).strip())
    fmts = [
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y %I:%M:%S %p", "%d/%m/%Y %I:%M %p",
        "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d-%m-%Y %I:%M:%S %p", "%d-%m-%Y %I:%M %p",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%d %b %Y %H:%M:%S", "%d %b %Y %H:%M",
        "%d %b %Y %I:%M:%S %p", "%d %b %Y %I:%M %p"
    ]
    for f in fmts:
        try: return datetime.strptime(s, f)
        except ValueError: pass
    try: return datetime.fromisoformat(s)
    except ValueError: return None

def parse_excel_analytics(file_bytes, password=""):
    """Parse an attendance or swipe log Excel spreadsheet (supports password-protected files) to produce full data analytics."""
    wb = load_excel_workbook(file_bytes, password)
    ws = wb.active

    # Find Header row and column indexes
    emp_col, name_col, swipes_col, type_col, lat_col, lng_col, area_col, mode_col, header_row = None, None, None, None, None, None, None, None, None
    
    for row in ws.iter_rows(max_row=25):
        for cell in row:
            if cell.value:
                val = str(cell.value).strip().lower()
                if val == "employee code":
                    emp_col = cell.column
                    header_row = cell.row
        if emp_col:
            break

    if not emp_col or not header_row:
        return None

    # Map other parameters on header row
    for c in range(1, ws.max_column + 1):
        cell_val = str(ws.cell(header_row, c).value or "").strip().lower()
        if cell_val == "employee name":
            name_col = c
        elif cell_val == "swipes":
            swipes_col = c
        elif cell_val == "type":
            type_col = c
        elif "latitude" in cell_val:
            lat_col = c
        elif "longit" in cell_val:
            lng_col = c
        elif cell_val == "area":
            area_col = c
        elif "attendance mode" in cell_val or cell_val == "mode":
            mode_col = c

    # Case A: Swipe Log format (contains 'Swipes' column)
    if swipes_col:
        employees = set()
        emp_names = {}
        records = []
        daily_counts = {}
        month_counts = {}
        emp_punch_count = {}
        area_counts = {}
        mode_counts = {}
        type_counts = {}

        for r in range(header_row + 1, ws.max_row + 1):
            emp_val = ws.cell(r, emp_col).value
            if not emp_val:
                continue
            emp_code = str(emp_val).strip().upper()
            if isinstance(emp_val, float) and emp_val.is_integer():
                emp_code = str(int(emp_val))
            
            employees.add(emp_code)
            emp_punch_count.setdefault(emp_code, 0)

            emp_name = str(ws.cell(r, name_col).value or "").strip() if name_col else ""
            if emp_name and emp_code not in emp_names:
                emp_names[emp_code] = emp_name

            raw_swipe = ws.cell(r, swipes_col).value
            dt = parse_swipe_dt(raw_swipe)
            if not dt:
                continue

            emp_punch_count[emp_code] += 1
            date_label = dt.strftime("%d-%b")
            month_key = dt.strftime("%b")
            time_display = dt.strftime("%H:%M:%S")

            daily_counts.setdefault(date_label, {"punches": 0, "total": len(employees), "times": []})
            daily_counts[date_label]["punches"] += 1
            daily_counts[date_label]["times"].append(time_display)

        # Recalculate accurate month_counts totals
        month_dates = {}
        for d_label in daily_counts:
            parts = d_label.split("-")
            m_key = parts[1].capitalize() if len(parts) >= 2 else d_label
            month_dates.setdefault(m_key, set())
            month_dates[m_key].add(d_label)

        num_employees = max(len(employees), 1)
        month_counts = {}
        for r_item in records:
            m_key = r_item["month"]
            month_counts.setdefault(m_key, {"punches": 0, "total": 0})
            month_counts[m_key]["punches"] += 1

        for m_key in month_counts:
            dates_in_m = len(month_dates.get(m_key, set()))
            expected_m = num_employees * max(dates_in_m, 1)
            month_counts[m_key]["total"] = expected_m
            month_counts[m_key]["dates_count"] = dates_in_m
            month_counts[m_key]["rate"] = min(100.0, round((month_counts[m_key]["punches"] / expected_m * 100), 1))

        total_dates = max(len(daily_counts), 1)
        total_possible = num_employees * total_dates
        total_actual_punches = len(records)
        completion_rate = round((total_actual_punches / total_possible * 100), 1) if total_possible > 0 else 100.0

        return {
            "is_swipe_log": True,
            "total_employees": num_employees,
            "total_dates": total_dates,
            "total_possible": total_possible,
            "total_punches": total_actual_punches,
            "completion_rate": completion_rate,
            "daily_counts": daily_counts,
            "month_counts": month_counts,
            "emp_punch_count": emp_punch_count,
            "emp_names": emp_names,
            "area_counts": area_counts,
            "mode_counts": mode_counts,
            "type_counts": type_counts,
            "records": records
        }


    # Case B: Attendance Sheet format (columns with dates or 'Punchin Time')
    punchin_cols = []
    for c in range(1, ws.max_column + 1):
        val = str(ws.cell(header_row, c).value or "").strip()
        if not val:
            continue
        if "punchin time" in val.lower():
            raw_date = val.lower().replace("punchin time", "").strip().capitalize()
            punchin_cols.append((c, raw_date, val))
        else:
            dt = parse_date_header(val)
            if dt:
                raw_date = dt.strftime("%d-%b")
                punchin_cols.append((c, raw_date, val))

    if not punchin_cols:
        return None

    employees = set()
    emp_names = {}
    records = []
    daily_counts = {}
    month_counts = {}
    emp_punch_count = {}

    for r in range(header_row + 1, ws.max_row + 1):
        emp_val = ws.cell(r, emp_col).value
        if not emp_val:
            continue
        emp_code = str(emp_val).strip().upper()
        if isinstance(emp_val, float) and emp_val.is_integer():
            emp_code = str(int(emp_val))
        employees.add(emp_code)
        emp_punch_count.setdefault(emp_code, 0)

        emp_name = str(ws.cell(r, name_col).value or "").strip() if name_col else ""
        if emp_name and emp_code not in emp_names:
            emp_names[emp_code] = emp_name

        for col_idx, date_label, col_name in punchin_cols:
            pval = ws.cell(r, col_idx).value
            daily_counts.setdefault(date_label, {"punches": 0, "total": 0, "times": []})
            daily_counts[date_label]["total"] += 1

            parts = date_label.split("-")
            month_key = parts[1].capitalize() if len(parts) >= 2 else date_label

            month_counts.setdefault(month_key, {"punches": 0, "total": 0})
            month_counts[month_key]["total"] += 1

            if pval is not None and str(pval).strip() != "":
                daily_counts[date_label]["punches"] += 1
                month_counts[month_key]["punches"] += 1
                emp_punch_count[emp_code] += 1
                
                time_display = str(pval)
                if isinstance(pval, (datetime, date)):
                    time_display = pval.strftime("%H:%M:%S")
                
                daily_counts[date_label]["times"].append(time_display)
                records.append({
                    "emp": emp_code,
                    "emp_name": emp_name or emp_code,
                    "date": date_label,
                    "month": month_key,
                    "time": time_display
                })

    total_possible = len(employees) * len(punchin_cols)
    total_actual_punches = len(records)
    completion_rate = round((total_actual_punches / total_possible * 100), 1) if total_possible > 0 else 0

    return {
        "is_swipe_log": False,
        "total_employees": len(employees),
        "total_dates": len(punchin_cols),
        "total_possible": total_possible,
        "total_punches": total_actual_punches,
        "completion_rate": completion_rate,
        "daily_counts": daily_counts,
        "month_counts": month_counts,
        "emp_punch_count": emp_punch_count,
        "emp_names": emp_names,
        "records": records
    }

def fetch_reports(token: str):
    """Fetch all processed report batches from the backend API."""
    try:
        r = requests.get(f"{API}/reports", headers={"Authorization": f"Bearer {token}"}, timeout=10)
        if r.ok:
            return r.json()
    except Exception:
        pass
    return []


def fetch_report_detail(job_uuid: str, token: str):
    """Fetch full analytics for a specific report job from the backend API."""
    try:
        r = requests.get(f"{API}/reports/{job_uuid}", headers={"Authorization": f"Bearer {token}"}, timeout=20)
        if r.ok:
            return r.json()
    except Exception:
        pass
    return None


def render_login():
    """Render the Attendance Report Record split-screen login page matching Image 2 UI design."""
    logo_data = get_asset_base64("iiit_logo.jpg")
    background_data = get_asset_base64("home_bg.jpg")

    background_style = (
        f"background-image: url('{background_data}');"
        if background_data
        else "background: linear-gradient(135deg, #00b4d8 0%, #0077b6 100%);"
    )

    logo_html = (
        f'<img src="{logo_data}" alt="IIIT Bangalore Logo" class="logo">'
        if logo_data
        else '<div class="logo-fallback" style="font-size: 24px; font-weight: bold; color: white;">IIIT-B</div>'
    )

    social_html = """
    <div class="social-icons">
        <a href="https://www.facebook.com/IIITBofficial" target="_blank"
           rel="noopener noreferrer" class="social-icon" aria-label="Facebook">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
                 stroke-linecap="round" stroke-linejoin="round">
                <path d="M18 2h-3a5 5 0 0 0-5 5v3H7v4h3v8h4v-8h3l1-4h-4V7a1 1 0 0 1 1-1h3z"/>
            </svg>
        </a>
        <a href="https://x.com/IIITB_official" target="_blank"
           rel="noopener noreferrer" class="social-icon" aria-label="X">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
                 stroke-linecap="round" stroke-linejoin="round">
                <path d="M18 6 6 18"/>
                <path d="m6 6 12 12"/>
            </svg>
        </a>
        <a href="https://www.linkedin.com/school/iiitbofficial" target="_blank"
           rel="noopener noreferrer" class="social-icon" aria-label="LinkedIn">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
                 stroke-linecap="round" stroke-linejoin="round">
                <path d="M16 8a6 6 0 0 1 6 6v7h-4v-7a2 2 0 0 0-4 0v7h-4v-7a6 6 0 0 1 6-6z"/>
                <rect width="4" height="12" x="2" y="9"/>
                <circle cx="4" cy="4" r="2"/>
            </svg>
        </a>
        <a href="https://www.youtube.com/user/iiitbmedia" target="_blank"
           rel="noopener noreferrer" class="social-icon" aria-label="YouTube">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
                 stroke-linecap="round" stroke-linejoin="round">
                <path d="M2.5 12s0-4.5.6-6.3a2.9 2.9 0 0 1 2-2C6.9 3.1 12 3.1 12 3.1s5.1 0 6.9.6a2.9 2.9 0 0 1 2 2c.6 1.8.6 6.3.6 6.3s0 4.5-.6 6.3a2.9 2.9 0 0 1-2 2c-1.8.6-6.9.6-6.9.6s-5.1 0-6.9-.6a2.9 2.9 0 0 1-2-2C2.5 16.5 2.5 12 2.5 12z"/>
                <path d="m10 8 5 4-5 4z"/>
            </svg>
        </a>
    </div>
    """

    col_left, col_right = st.columns([7.5, 2.5])

    with col_left:
        st.markdown(
            clean_html(f"""
            <div class="branding-section" style="{background_style}">
                <div class="branding-content">
                    <div class="logo-container">
                        {logo_html}
                    </div>

                    <div class="institution-info">
                        <div class="institution-name">International Institute of</div>
                        <div class="institution-name">
                            Information Technology Bangalore
                        </div>
                    </div>

                    <div class="portal-title">
                        <div class="portal-title-text">Attendance Record PORTAL</div>
                    </div>
                </div>

                <div class="branding-footer">
                    {social_html}
                    <div class="copyright">
                        <p>© 2026 International Institute of Information Technology - Bangalore</p>
                        <p class="support-text">Technical Support - application@iiitb.ac.in</p>
                    </div>
                </div>
            </div>
            """),
            unsafe_allow_html=True,
        )

    with col_right:
        st.markdown(
            clean_html("""
            <div style="text-align: center; margin-bottom: 30px; margin-top: 20px;">
                <h1 style="color: #0077ff; font-size: 34px; font-weight: 700; margin: 0; font-family: 'Inter', sans-serif;">Login</h1>
            </div>
            """),
            unsafe_allow_html=True,
        )

        with st.form("login_form", clear_on_submit=False):
            u = st.text_input(
                "Username",
                placeholder="Enter Username",
                label_visibility="collapsed",
                key="login_username",
            )

            st.markdown('<div style="height: 18px;"></div>', unsafe_allow_html=True)

            p = st.text_input(
                "Password",
                type="password",
                placeholder="Enter Password",
                label_visibility="collapsed",
                key="login_password",
            )

            st.markdown('<div style="height: 10px;"></div>', unsafe_allow_html=True)

            submitted = st.form_submit_button(
                "Login",
                type="primary",
                use_container_width=True,
            )

            if submitted:
                if not u or not p:
                    st.error("Please enter both username and password.")
                else:
                    try:
                        r = requests.post(
                            f"{API}/auth/login",
                            json={"username": u, "password": p},
                            timeout=15,
                        )

                        if r.ok:
                            res_data = r.json()
                            st.session_state.token = res_data["access_token"]
                            st.session_state.user = res_data.get("username", u)
                            st.session_state.role = res_data.get(
                                "role", "coordinator"
                            )
                            st.success(
                                f"Signed in successfully as {st.session_state.role}!"
                            )
                            st.rerun()
                        else:
                            try:
                                detail = r.json().get("detail", "Login failed.")
                            except ValueError:
                                detail = "Login failed."
                            st.error(detail)

                    except requests.RequestException:
                        st.error("Cannot connect to FastAPI backend service.")


def render_navbar():
    col_brand, col_nav, col_actions = st.columns([3, 5, 4])
    
    with col_brand:
        st.markdown(clean_html("""
            <div style="display: flex; align-items: center; gap: 0.5rem; padding-top: 0.2rem;">
                <span style="font-size: 1.5rem;">⏱️</span>
                <div style="font-weight: 700; font-size: 1.15rem; color: #f8fafc;">Attendance Punch-In</div>
            </div>
        """), unsafe_allow_html=True)
        
    with col_nav:
        n1, n2, n3 = st.columns(3)
        with n1:
            if st.button("📊 Dashboard", type="primary" if st.session_state.active_tab == "dashboard" else "secondary", use_container_width=True):
                st.session_state.active_tab = "dashboard"
                st.rerun()
        with n2:
            if st.button("📤 Upload & Process", type="primary" if st.session_state.active_tab == "upload" else "secondary", use_container_width=True):
                st.session_state.active_tab = "upload"
                st.rerun()
        with n3:
            if st.button("📈 Reports", type="primary" if st.session_state.active_tab == "report" else "secondary", use_container_width=True):
                st.session_state.active_tab = "report"
                st.rerun()

    with col_actions:
        # User & Role badge placed right next to the logout button
        u_col, btn_col = st.columns([2.2, 1.2])
        role_icon = "🎓" if st.session_state.get("role") == "dean-faculty" else "👤"
        role_style = "background: rgba(16, 185, 129, 0.15); border-color: rgba(16, 185, 129, 0.3); color: #34d399;" if st.session_state.get("role") == "dean-faculty" else ""
        
        with u_col:
            st.markdown(clean_html(f"""
                <div style="text-align: right; padding-top: 0.35rem;">
                    <span class="user-badge" style="{role_style}">{role_icon} {st.session_state.get('user', 'user')}</span>
                </div>
            """), unsafe_allow_html=True)
        with btn_col:
            if st.button("Logout", key="top_logout", type="secondary", use_container_width=True):
                st.session_state.clear()
                st.rerun()

    st.divider()


def create_plotly_bar_chart(x_vals, y_vals, title, x_title, y_title, is_percentage=False):
    """Create a high-end, responsive Plotly dark-themed bar chart."""
    fig = go.Figure()
    
    hover_template = "%{x}: <b>%{y:.1f}%</b><extra></extra>" if is_percentage else "%{x}: <b>%{y:,}</b><extra></extra>"
    text_labels = [f"{v:.1f}%" if is_percentage else f"{v:,}" for v in y_vals]

    fig.add_trace(go.Bar(
        x=x_vals,
        y=y_vals,
        marker=dict(
            color=y_vals,
            colorscale=[[0, "#0077b6"], [0.5, "#00b4d8"], [1.0, "#38bdf8"]],
            line=dict(color="rgba(255,255,255,0.2)", width=1)
        ),
        hovertemplate=hover_template,
        text=text_labels,
        textposition="outside",
        textfont=dict(color="#f8fafc", size=11, family="Inter, sans-serif")
    ))
    
    fig.update_layout(
        title=dict(text=title, font=dict(size=15, color="#f8fafc", family="Inter, sans-serif")),
        xaxis=dict(title=x_title, tickfont=dict(color="#cbd5e1"), gridcolor="rgba(255,255,255,0.05)"),
        yaxis=dict(title=y_title, tickfont=dict(color="#cbd5e1"), gridcolor="rgba(255,255,255,0.08)"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15, 23, 42, 0.6)",
        margin=dict(l=20, r=20, t=40, b=30),
        height=340,
        showlegend=False
    )
    return fig

def create_plotly_area_chart(x_vals, y_vals, title, x_title, y_title):
    """Create a high-end Plotly dark-themed area line trend chart."""
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=x_vals,
        y=y_vals,
        mode="lines+markers",
        fill="tozeroy",
        fillcolor="rgba(0, 180, 216, 0.15)",
        line=dict(color="#00b4d8", width=3, shape="spline"),
        marker=dict(size=7, color="#0077b6", line=dict(color="#ffffff", width=1.5)),
        hovertemplate="%{x}: <b>%{y:,} Swipes</b><extra></extra>"
    ))
    
    fig.update_layout(
        title=dict(text=title, font=dict(size=15, color="#f8fafc", family="Inter, sans-serif")),
        xaxis=dict(title=x_title, tickfont=dict(color="#cbd5e1"), gridcolor="rgba(255,255,255,0.05)"),
        yaxis=dict(title=y_title, tickfont=dict(color="#cbd5e1"), gridcolor="rgba(255,255,255,0.08)"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15, 23, 42, 0.6)",
        margin=dict(l=20, r=20, t=40, b=30),
        height=340
    )
    return fig

def create_plotly_donut_chart(labels, values, title):
    """Create a high-end Plotly donut chart for mode & area breakdowns."""
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.55,
        marker=dict(colors=["#0077b6", "#00b4d8", "#38bdf8", "#34d399", "#f59e0b", "#ec4899"]),
        textinfo="percent+label",
        hoverinfo="label+value+percent",
        textfont=dict(color="#ffffff")
    )])
    
    fig.update_layout(
        title=dict(text=title, font=dict(size=15, color="#f8fafc", family="Inter, sans-serif")),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=40, b=20),
        height=340,
        showlegend=True,
        legend=dict(font=dict(color="#cbd5e1"))
    )
    return fig

def render_dashboard():
    st.markdown("## 📊 Attendance Analytics Dashboard")
    st.caption("Select a processed report cycle below to view monthwise, datewise, and employee analytics.")

    # Refresh button
    _, col_refresh = st.columns([6, 1])
    with col_refresh:
        if st.button("🔄 Refresh", key="refresh_reports_btn", use_container_width=True):
            st.session_state.pop("available_reports", None)
            st.session_state.pop("selected_report_uuid", None)
            st.session_state.analytics_data = None

    # Load available reports (lazily cached in session_state)
    if st.session_state.get("available_reports") is None:
        with st.spinner("⏳ Loading available reports from database..."):
            st.session_state.available_reports = fetch_reports(st.session_state.token)

    reports = st.session_state.get("available_reports") or []

    if not reports:
        # ─── Empty State ────────────────────────────────────────────────────
        st.markdown(clean_html("""
            <div class="glass-card" style="text-align: center; padding: 3rem 2rem; margin-top: 1rem;">
                <div style="font-size: 3.5rem; margin-bottom: 1rem;">📭</div>
                <h3 style="color: #94a3b8; margin-bottom: 0.5rem;">No Reports Generated Yet</h3>
                <p style="color: #64748b; font-size: 0.95rem; max-width: 480px; margin: 0 auto;">
                    Process your Attendance &amp; Swipe Excel files first to generate analytics reports.
                    Once processed, all report cycles will appear here automatically.
                </p>
            </div>
        """), unsafe_allow_html=True)

        c1, c2, c3, c4 = st.columns(4)
        for col, val, label in [
            (c1, "—", "Total Employees"), (c2, "—", "Dates Analyzed"),
            (c3, "—%", "Punch-In Rate"), (c4, "🟢", "System Status")
        ]:
            with col:
                st.markdown(f'<div class="stat-box"><div class="stat-val">{val}</div><div class="stat-label">{label}</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📤 Go to Upload & Process →", type="primary"):
            st.session_state.active_tab = "upload"
            st.rerun()
        return

    # ─── Report Cycle Selector ──────────────────────────────────────────────
    options = {}
    for rep in reports:
        cycle = rep.get("cycle_label", "Unknown Cycle")
        fname = (rep.get("attendance_filename") or "Unnamed").replace(".xlsx", "").replace(".xlsm", "")
        created = (rep.get("created_at") or "")[:10]
        label = f"📅  {cycle}   |   {fname}   ({created})"
        options[label] = rep["job_uuid"]

    default_idx = 0
    if st.session_state.get("selected_report_uuid") in list(options.values()):
        default_idx = list(options.values()).index(st.session_state.selected_report_uuid)

    selected_label = st.selectbox(
        "📂 Select Report Cycle to Analyze",
        list(options.keys()),
        index=default_idx,
        key="report_cycle_selector"
    )
    selected_uuid = options[selected_label]

    # Load analytics for the selected report if changed or not yet loaded
    if st.session_state.get("selected_report_uuid") != selected_uuid or st.session_state.analytics_data is None:
        with st.spinner("⏳ Loading report analytics from database..."):
            detail = fetch_report_detail(selected_uuid, st.session_state.token)
            if detail:
                st.session_state.analytics_data = detail
                st.session_state.analytics_file_name = detail.get("attendance_filename", "Report")
                st.session_state.selected_report_uuid = selected_uuid
            else:
                st.error("❌ Could not load analytics for this report. It may have no attendance records stored yet.")
                return

    data = st.session_state.analytics_data
    if not data:
        st.warning("No analytics data loaded. Please select a report above.")
        return

    # ─── Report Header ──────────────────────────────────────────────────────
    cycle_label = data.get("cycle_label", "Report")
    st.markdown(f"### 📈 Analytics: `{cycle_label}`")

    # ─── 4 Key Metric Cards ──────────────────────────────────────────────────
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(f'<div class="stat-box"><div class="stat-val">{data["total_employees"]}</div><div class="stat-label">Total Employees</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="stat-box"><div class="stat-val">{data["total_dates"]}</div><div class="stat-label">Dates Analyzed</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="stat-box"><div class="stat-val">{data["total_punches"]:,} / {data["total_possible"]:,}</div><div class="stat-label">Matched Punch-Ins</div></div>', unsafe_allow_html=True)
    with m4:
        st.markdown(f'<div class="stat-box"><div class="stat-val" style="color: #34d399;">{data["completion_rate"]}%</div><div class="stat-label">Punch-In Completion Rate</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ─── MONTHWISE ANALYTICS ─────────────────────────────────────────────────
    st.markdown("### 🗓️ Month-Wise Analytics")
    month_counts = data.get("month_counts", {})
    col_m1, col_m2 = st.columns([1, 1])
    with col_m1:
        st.markdown("#### Summary Table")
        month_table_data = []
        for m_name, m_info in month_counts.items():
            rate_val = m_info.get("rate", round((m_info["punches"] / m_info["total"] * 100), 1) if m_info.get("total", 0) > 0 else 0)
            month_table_data.append({"Month": m_name, "Punch-Ins Recorded": m_info["punches"], "Expected Slots": m_info["total"], "Completion Rate": f"{rate_val}%"})
        st.dataframe(month_table_data, use_container_width=True)
    with col_m2:
        if month_counts:
            month_names = list(month_counts.keys())
            month_rates = [m_info.get("rate", round((m_info["punches"] / m_info["total"] * 100), 1) if m_info.get("total", 0) > 0 else 0) for m_info in month_counts.values()]
            fig_month = create_plotly_bar_chart(month_names, month_rates, "Monthwise Completion Rate (%)", "Month", "Completion Rate (%)", is_percentage=True)
            st.plotly_chart(fig_month, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ─── DATEWISE ANALYTICS ──────────────────────────────────────────────────
    st.markdown("### 📅 Date-Wise Analytics")
    daily_counts = data.get("daily_counts", {})
    if daily_counts:
        date_labels = list(daily_counts.keys())
        daily_punches = [d_info["punches"] for d_info in daily_counts.values()]
        fig_date = create_plotly_area_chart(date_labels, daily_punches, "Daily Punch-In Volume Trend", "Date", "Punch-Ins")
        st.plotly_chart(fig_date, use_container_width=True)

    st.markdown("#### 📋 Date-Wise Breakdown (Daily Summary)")
    daily_table = []
    for d_label, d_info in daily_counts.items():
        punches = d_info["punches"]
        tot = d_info["total"]
        pct = round((punches / tot * 100), 1) if tot > 0 else 0
        # Prefer explicit first/last punch stored by API; fall back to min/max of times list
        times = d_info.get("times", [])
        first_punch = d_info.get("first_punch") or (min(times) if times else "—")
        last_punch  = d_info.get("last_punch")  or (max(times) if times else "—")
        daily_table.append({
            "Date": d_label,
            "Punch-Ins": punches,
            "Total Employees": tot,
            "Attendance %": f"{pct}%",
            "Earliest Punch-In": first_punch or "—",
            "Latest Punch-In": last_punch or "—"
        })
    st.dataframe(daily_table, use_container_width=True)

    # Per-employee per-date detail table
    emp_daily_records = data.get("employee_daily_records", [])
    if emp_daily_records:
        st.markdown("#### 👤 Employee-Wise Punch-In Detail")
        emp_names_map = data.get("emp_names", {})
        emp_detail_table = []
        for row in emp_daily_records:
            emp_code = row.get("employee_code", "")
            emp_name = row.get("employee_name") or emp_names_map.get(emp_code, "")
            status = row.get("status", "—")
            status_icon = "✅ Present" if status == "Present" else "❌ Absent"
            emp_detail_table.append({
                "Date": row.get("date", "—"),
                "Employee Code": emp_code,
                "Employee Name": emp_name or "—",
                "Punch-In Time": row.get("punch_in_time", "—"),
                "Status": status_icon
            })
        st.dataframe(emp_detail_table, use_container_width=True)


    st.markdown("<br>", unsafe_allow_html=True)

    # ─── EMPLOYEE INSIGHTS ───────────────────────────────────────────────────
    st.markdown("### 👤 Employee Attendance Insights")
    emp_counts = data.get("emp_punch_count", {})
    emp_names  = data.get("emp_names", {})
    total_dates = data.get("total_dates", 1) or 1

    ec1, ec2 = st.columns([1, 1])
    with ec1:
        st.markdown("#### 🏆 Top Performers (Most Punch-Ins)")
        sorted_emps = sorted(emp_counts.items(), key=lambda x: x[1], reverse=True)[:10]
        top_emp_data = []
        for emp, count in sorted_emps:
            row = {"Employee Code": emp, "Total Punch-Ins": count, "Attendance %": f"{round(count / total_dates * 100, 1)}%"}
            if emp_names.get(emp):
                row["Employee Name"] = emp_names[emp]
            top_emp_data.append(row)
        st.dataframe(top_emp_data, use_container_width=True)

    with ec2:
        st.markdown("#### ⚠️ Employees with Missing Punch-Ins")
        missing_emps = []
        for emp, count in emp_counts.items():
            if count < total_dates:
                row = {"Employee Code": emp, "Punches Found": count, "Missing Days": total_dates - count, "Attendance %": f"{round(count / total_dates * 100, 1)}%"}
                if emp_names.get(emp):
                    row["Employee Name"] = emp_names[emp]
                missing_emps.append(row)
        missing_emps.sort(key=lambda x: x["Missing Days"], reverse=True)
        if missing_emps:
            st.dataframe(missing_emps[:10], use_container_width=True)
        else:
            st.success("🎉 All employees have complete punch-in records!")


def render_upload():
    st.markdown("## 📤 Upload & Process Attendance")
    st.caption("Upload Attendance Master file and Swipe file to automatically insert Punchin Times.")

    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(clean_html("""
            <div class="glass-card">
                <h4 style="color: #00b4d8; margin: 0;">1. Attendance Master Excel</h4>
            </div>
        """), unsafe_allow_html=True)
        af = st.file_uploader("Select Attendance File", type=["xlsx", "xlsm"], key="att_file")
        ap = st.text_input("Attendance Password (if encrypted)", type="password", placeholder="Leave blank if not password protected", key="att_pass")

    with col2:
        st.markdown(clean_html("""
            <div class="glass-card">
                <h4 style="color: #38bdf8; margin: 0;">2. Swipe Log Excel</h4>
            </div>
        """), unsafe_allow_html=True)
        sf = st.file_uploader("Select Swipe Log File", type=["xlsx", "xlsm"], key="swp_file")
        sp = st.text_input("Swipe Log Password (if encrypted)", type="password", placeholder="Leave blank if not password protected", key="swp_pass")

    st.markdown("<br>", unsafe_allow_html=True)
    
    st.info("💡 **Rule**: Employee Code + Date are matched across files. The earliest swipe timestamp is picked as Punchin Time. Dates without swipes remain blank.")

    if st.button("⚡ Generate & Process Punch-In Excel", type="primary", use_container_width=True):
        if not af or not sf:
            st.error("⚠️ Please select both the Attendance Excel file and Swipe Log Excel file.")
            return
        
        with st.spinner("⏳ Processing Excel spreadsheets, parsing swipes, and calculating punch-in times..."):
            try:
                r = requests.post(
                    f"{API}/process",
                    headers={"Authorization": f"Bearer {st.session_state.token}"},
                    files={
                        "attendance_file": (af.name, af.getvalue(), af.type),
                        "swipe_file": (sf.name, sf.getvalue(), sf.type)
                    },
                    data={
                        "attendance_password": ap,
                        "swipe_password": sp
                    },
                    timeout=300
                )
                
                if r.ok:
                    valid_swipes = int(r.headers.get("X-Valid-Swipes", "0"))
                    invalid_swipes = int(r.headers.get("X-Invalid-Swipes", "0"))
                    date_count = int(r.headers.get("X-Date-Count", "0"))
                    
                    st.success("✅ Updated Attendance Excel generated successfully!")

                    # Clear reports cache so Dashboard auto-loads the new report on next visit
                    st.session_state.pop("available_reports", None)
                    st.session_state.pop("selected_report_uuid", None)
                    st.session_state.analytics_data = None

                    # Display processing stats badges
                    sc1, sc2, sc3 = st.columns(3)
                    with sc1:
                        st.metric("Valid Swipes Processed", f"{valid_swipes:,}")
                    with sc2:
                        st.metric("Invalid / Skipped Entries", f"{invalid_swipes:,}")
                    with sc3:
                        st.metric("Date Columns Updated", f"{date_count}")

                    # Save job to session history
                    st.session_state.history.append({
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "attendance_file": af.name,
                        "swipe_file": sf.name,
                        "valid_swipes": valid_swipes,
                        "invalid_swipes": invalid_swipes,
                        "date_count": date_count
                    })

                    st.download_button(
                        label="📥 Download Updated Attendance Excel",
                        data=BytesIO(r.content),
                        file_name="Updated_Attendance_Punchin.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )
                elif r.status_code == 401:
                    st.session_state.clear()
                    st.error("Session expired. Please sign in again.")
                    st.rerun()
                else:
                    detail = r.json().get("detail", "Processing failed.")
                    st.error(f"❌ Error: {detail}")
            except requests.RequestException as e:
                st.error(f"❌ Backend connection error: {e}")

def render_reports():
    st.markdown("## 📈 Processing Reports & Analytics Log")
    st.caption("Review detailed records, batch logs, and swipe match statistics.")

    if not st.session_state.history:
        st.info("ℹ️ No processing batches run in this session yet. Upload files in **Upload & Process** to generate reports!")
        return

    # Aggregate Analytics
    total_runs = len(st.session_state.history)
    total_valid = sum(h["valid_swipes"] for h in st.session_state.history)
    total_invalid = sum(h["invalid_swipes"] for h in st.session_state.history)
    
    rc1, rc2, rc3 = st.columns(3)
    with rc1:
        st.metric("Total Batches Processed", total_runs)
    with rc2:
        st.metric("Cumulative Valid Swipes", f"{total_valid:,}")
    with rc3:
        st.metric("Cumulative Invalid Swipes", f"{total_invalid:,}")

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📜 Batch Execution Logs")
    
    st.dataframe(
        st.session_state.history,
        use_container_width=True,
        column_config={
            "timestamp": "Execution Time",
            "attendance_file": "Attendance File",
            "swipe_file": "Swipe File",
            "valid_swipes": "Valid Swipes",
            "invalid_swipes": "Invalid / Skipped",
            "date_count": "Dates Updated"
        }
    )

def main():
    inject_custom_css(login_page=not bool(st.session_state.token))

    if not st.session_state.token:
        render_login()
        return

    render_navbar()

    # Render selected page tab
    if st.session_state.active_tab == "dashboard":
        render_dashboard()
    elif st.session_state.active_tab == "upload":
        render_upload()
    elif st.session_state.active_tab == "report":
        render_reports()

if __name__ == "__main__":
    main()