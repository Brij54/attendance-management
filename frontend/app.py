import os, requests, re, tempfile, msoffcrypto, base64
from textwrap import dedent
from io import BytesIO
from pathlib import Path
from datetime import datetime, date
from dotenv import load_dotenv
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook
import plotly.express as px
import plotly.graph_objects as go

# Create assets folder if not exists
ASSETS_DIR = Path(__file__).resolve().parent / "assets"
ASSETS_DIR.mkdir(parents=True, exist_ok=True)

_ASSET_CACHE = {}

def get_asset_base64(filename):
    """Retrieve an asset file (logo, background) as Base64 data URL if present.
    Caches result in-memory to prevent expensive disk I/O on every Streamlit rerun.
    """
    if filename in _ASSET_CACHE:
        return _ASSET_CACHE[filename]

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
                    res = f"data:{mime};base64,{base64.b64encode(data).decode()}"
                    _ASSET_CACHE[filename] = res
                    return res
                except Exception:
                    pass
    return None


# Load environment variables
load_dotenv(Path(__file__).resolve().parents[1] / ".env.production")
load_dotenv(Path(__file__).resolve().parent / ".env.production")
load_dotenv(Path(__file__).resolve().parents[1] / ".env")
load_dotenv(Path(__file__).resolve().parent / ".env")
load_dotenv()

st.set_page_config(
    page_title="IIIT-B Attendance Portal",
    page_icon=str(ASSETS_DIR / "IIITB_logo1.png") if (ASSETS_DIR / "IIITB_logo1.png").exists() else None,
    layout="wide",
    initial_sidebar_state="expanded"
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
if "sidebar_open" not in st.session_state:
    st.session_state.sidebar_open = True




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
        header[data-testid="stHeader"],
        .stAppHeader,
        .stAppDeployButton,
        div[data-testid="stToolbar"],
        #MainMenu,
        [data-testid="stDecoration"],
        footer,
        [data-testid="stSidebar"] {
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
        sidebar_is_open = st.session_state.get("sidebar_open", True)
        sidebar_css = """
        [data-testid="stSidebar"],
        section[data-testid="stSidebar"] {
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
            transform: none !important;
            margin-left: 0 !important;
            width: 250px !important;
            min-width: 250px !important;
            max-width: 250px !important;
            background: #ffffff !important;
            border-right: 1px solid #e2e8f0 !important;
            box-shadow: 3px 0 12px rgba(0, 0, 0, 0.05) !important;
            padding-top: 0 !important;
            margin-top: 0 !important;
            transition: width 0.15s ease-out, min-width 0.15s ease-out, max-width 0.15s ease-out !important;
        }

        [data-testid="stSidebar"] .block-container {
            padding: 0.35rem 0.85rem 1.25rem 0.85rem !important;
        }

        [data-testid="stSidebar"] .stButton > button {
            height: 44px !important;
            text-align: left !important;
            justify-content: flex-start !important;
            padding: 0 1rem !important;
            font-size: 0.92rem !important;
            font-weight: 600 !important;
            border-radius: 8px !important;
            margin-bottom: 0.35rem !important;
            display: flex !important;
            align-items: center !important;
            gap: 0.65rem !important;
            width: 100% !important;
        }
        """ if sidebar_is_open else """
        [data-testid="stSidebar"],
        section[data-testid="stSidebar"] {
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
            transform: none !important;
            margin-left: 0 !important;
            width: 68px !important;
            min-width: 68px !important;
            max-width: 68px !important;
            background: #ffffff !important;
            border-right: 1px solid #e2e8f0 !important;
            box-shadow: 2px 0 8px rgba(0, 0, 0, 0.04) !important;
            overflow: hidden !important;
            padding-top: 0 !important;
            margin-top: 0 !important;
            transition: width 0.15s ease-out, min-width 0.15s ease-out, max-width 0.15s ease-out !important;
        }

        [data-testid="stSidebar"] .block-container {
            padding: 0.35rem 0.25rem 1.25rem 0.25rem !important;
        }

        [data-testid="stSidebar"] .stButton > button {
            width: 46px !important;
            min-width: 46px !important;
            max-width: 46px !important;
            height: 46px !important;
            padding: 0 !important;
            margin: 0.25rem auto !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            font-size: 1.25rem !important;
            border-radius: 8px !important;
        }
        """

        st.markdown(clean_html(f"""
        <style>
        /* Clean top: hide Streamlit's native header toolbar, deploy button, native sidebar header, and collapse chevrons */
        header[data-testid="stHeader"],
        .stAppHeader,
        .stAppDeployButton,
        div[data-testid="stToolbar"],
        #MainMenu,
        [data-testid="stDecoration"],
        footer,
        [data-testid="stSidebarHeader"],
        header[data-testid="stSidebarHeader"],
        div[data-testid="stSidebarHeader"],
        [data-testid="stSidebarCollapseButton"],
        [data-testid="stSidebarCollapsedControl"],
        [data-testid="collapsedControl"],
        [data-testid="stHeaderActionElements"],
        [data-testid="stHeadingWithActionElements"] a,
        a.anchorjs-link,
        a.anchor-link {{
            display: none !important;
            visibility: hidden !important;
            height: 0 !important;
            min-height: 0 !important;
            max-height: 0 !important;
            padding: 0 !important;
            margin: 0 !important;
            opacity: 0 !important;
            pointer-events: none !important;
            overflow: hidden !important;
        }}

        /* Completely remove all top whitespace in sidebar containers */
        [data-testid="stSidebarUserContent"],
        [data-testid="stSidebarContent"],
        section[data-testid="stSidebar"] > div:first-child,
        [data-testid="stSidebar"] > div:first-child {{
            padding-top: 0 !important;
            margin-top: 0 !important;
        }}

        [data-testid="stSidebar"] div[data-testid="stVerticalBlock"]:first-child,
        [data-testid="stSidebar"] div[data-testid="stVerticalBlock"] > div:first-child,
        [data-testid="stSidebar"] [data-testid="stHorizontalBlock"]:first-child {{
            padding-top: 0 !important;
            margin-top: 0 !important;
        }}

        * {{
            box-sizing: border-box;
        }}

        html, body {{
            overflow-x: hidden !important;
            background-color: #f4f6f9 !important;
        }}

        .stApp {{
            background-color: #f4f6f9 !important;
            color: #1e293b !important;
            font-family: 'Inter', system-ui, -apple-system, sans-serif !important;
        }}

        /* Dynamic Sidebar Styling based on session state */
        {sidebar_css}

        [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {{
            gap: 0.2rem !important;
        }}

        /* Sidebar Navigation Buttons - Matching User's Screenshot */
        [data-testid="stSidebar"] .stButton > button {{
            height: 48px !important;
            text-align: left !important;
            justify-content: flex-start !important;
            padding: 0 1.1rem !important;
            font-size: 0.95rem !important;
            border-radius: 10px !important;
            margin-bottom: 0.35rem !important;
            display: flex !important;
            align-items: center !important;
            gap: 0.85rem !important;
            width: 100% !important;
            border: none !important;
            transition: all 0.15s ease-in-out !important;
        }}

        /* Active Nav Item: Light gray pill container with bold black text */
        [data-testid="stSidebar"] .stButton > button[kind="primary"] {{
            background: #f1f3f5 !important;
            background-color: #f1f3f5 !important;
            color: #000000 !important;
            font-weight: 700 !important;
            border: none !important;
            box-shadow: none !important;
        }}

        [data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {{
            background: #e9ecef !important;
            background-color: #e9ecef !important;
            color: #000000 !important;
        }}

        /* Inactive Nav Items: Transparent background with dark slate text */
        [data-testid="stSidebar"] .stButton > button[kind="secondary"] {{
            background: transparent !important;
            background-color: transparent !important;
            border: none !important;
            color: #1e293b !important;
            font-weight: 500 !important;
            box-shadow: none !important;
        }}

        [data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover {{
            background: #f8fafc !important;
            background-color: #f8fafc !important;
            color: #000000 !important;
        }}

        /* Sidebar Header Hamburger Toggle Button */
        [data-testid="stSidebar"] [data-testid="stColumn"]:first-child .stButton > button {{
            width: 44px !important;
            min-width: 44px !important;
            max-width: 44px !important;
            height: 44px !important;
            padding: 0 !important;
            font-size: 1.35rem !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            background: transparent !important;
            border: none !important;
            color: #1e293b !important;
            border-radius: 8px !important;
        }}

        [data-testid="stSidebar"] [data-testid="stColumn"]:first-child .stButton > button:hover {{
            background: #f1f5f9 !important;
            color: #000000 !important;
        }}

        [data-testid="stAppViewContainer"] {{
            display: flex !important;
            flex-direction: row !important;
            overflow-x: hidden !important;
        }}

        [data-testid="stAppViewContainer"] > section.main {{
            flex: 1 1 auto !important;
            width: 100% !important;
            min-width: 0 !important;
        }}

        section.main .block-container,
        [data-testid="stMainBlockContainer"],
        [data-testid="stAppViewContainer"] > section.main {{
            padding-top: 0.85rem !important;
            padding-bottom: 2.5rem !important;
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
            max-width: 100% !important;
            margin: 0 auto !important;
            height: auto !important;
            overflow: visible !important;
            background: #f4f6f9 !important;
        }}

        /* Top Teal App Bar matching reference screenshot */
        .top-teal-bar {{
            background: #009bbd;
            background: linear-gradient(90deg, #009bbd 0%, #00a8cc 100%);
            height: 56px;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 1.25rem;
            margin-bottom: 0.4rem;
            box-shadow: 0 2px 8px rgba(0, 155, 189, 0.2);
        }}

        .top-teal-title {{
            font-size: 1.35rem;
            font-weight: 700;
            color: #ffffff;
            letter-spacing: -0.01em;
        }}

        .top-teal-right {{
            display: flex;
            align-items: center;
            gap: 0.85rem;
        }}

        .user-name-label {{
            color: #ffffff;
            font-size: 0.9rem;
            font-weight: 600;
            text-transform: capitalize;
            letter-spacing: 0.02em;
        }}

        .user-avatar-badge {{
            width: 36px;
            height: 36px;
            border-radius: 50%;
            background: #f43f5e;
            color: #ffffff;
            font-weight: 700;
            font-size: 1rem;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 2px 6px rgba(244, 63, 94, 0.35);
        }}

        .top-teal-logout-btn {{
            background: rgba(255, 255, 255, 0.18) !important;
            color: #ffffff !important;
            border: 1px solid rgba(255, 255, 255, 0.45) !important;
            border-radius: 7px !important;
            padding: 6px 14px !important;
            font-size: 0.85rem !important;
            font-weight: 600 !important;
            text-decoration: none !important;
            display: inline-flex !important;
            align-items: center !important;
            gap: 6px !important;
            transition: all 0.2s ease !important;
            cursor: pointer !important;
        }}

        .top-teal-logout-btn:hover {{
            background: rgba(255, 255, 255, 0.3) !important;
            border-color: #ffffff !important;
            color: #ffffff !important;
            text-decoration: none !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15) !important;
        }}

        /* Clean Enterprise Light Cards */
        .glass-card {{
            background: #ffffff !important;
            border: 1px solid #e2e8f0 !important;
            border-radius: 10px !important;
            padding: 1.25rem !important;
            margin-bottom: 1.25rem !important;
            box-shadow: 0 1px 4px rgba(0, 0, 0, 0.04) !important;
            color: #1e293b !important;
        }}

        .glass-card:hover {{
            border-color: #cbd5e1 !important;
        }}

        /* Clean Enterprise Light Stat Box */
        .stat-box {{
            background: #ffffff !important;
            border: 1px solid #e2e8f0 !important;
            border-radius: 10px !important;
            padding: 1.25rem !important;
            text-align: center !important;
            box-shadow: 0 1px 4px rgba(0, 0, 0, 0.04) !important;
        }}

        .stat-val {{
            font-size: 1.85rem !important;
            font-weight: 700 !important;
            color: #009bbd !important;
        }}

        .stat-label {{
            font-size: 0.82rem !important;
            color: #64748b !important;
            text-transform: uppercase !important;
            letter-spacing: 0.05em !important;
            margin-top: 0.25rem !important;
            font-weight: 600 !important;
        }}

        .stButton > button {{
            border-radius: 8px !important;
            font-weight: 600 !important;
            transition: all 0.2s ease-in-out !important;
        }}

        .stButton > button[kind="primary"] {{
            background: #009bbd !important;
            border: none !important;
            color: #ffffff !important;
            box-shadow: 0 2px 6px rgba(0, 155, 189, 0.3) !important;
        }}

        .stButton > button[kind="primary"]:hover {{
            background: #0087a4 !important;
            box-shadow: 0 4px 10px rgba(0, 155, 189, 0.4) !important;
        }}

        .stButton > button[kind="secondary"] {{
            background: #ffffff !important;
            border: 1px solid #cbd5e1 !important;
            color: #334155 !important;
        }}

        .stButton > button[kind="secondary"]:hover {{
            background: #f1f5f9 !important;
            border-color: #94a3b8 !important;
            color: #0f172a !important;
        }}

        div[data-baseweb="input"] {{
            background-color: #ffffff !important;
            border: 1.5px solid #cbd5e1 !important;
            border-radius: 8px !important;
            color: #0f172a !important;
        }}

        div[data-baseweb="input"]:focus-within {{
            border-color: #009bbd !important;
            box-shadow: 0 0 0 3px rgba(0, 155, 189, 0.2) !important;
        }}

        div[data-baseweb="input"] input {{
            color: #0f172a !important;
        }}

        /* BaseWeb Select Box - Crisp Light Enterprise Theme */
        div[data-baseweb="select"] {{
            background-color: #ffffff !important;
            border-radius: 8px !important;
        }}

        div[data-baseweb="select"] > div {{
            background-color: #ffffff !important;
            border: 1.5px solid #cbd5e1 !important;
            border-radius: 8px !important;
            color: #0f172a !important;
            min-height: 42px !important;
        }}

        div[data-baseweb="select"] > div:hover {{
            border-color: #94a3b8 !important;
        }}

        div[data-baseweb="select"]:focus-within > div {{
            border-color: #009bbd !important;
            box-shadow: 0 0 0 3px rgba(0, 155, 189, 0.2) !important;
        }}

        div[data-baseweb="select"] div[role="combobox"] {{
            color: #0f172a !important;
            font-weight: 500 !important;
            background-color: #ffffff !important;
        }}

        div[data-baseweb="select"] span {{
            color: #0f172a !important;
            font-weight: 500 !important;
        }}

        div[data-baseweb="select"] svg {{
            fill: #475569 !important;
            color: #475569 !important;
        }}

        /* BaseWeb Select Dropdown Popover & Menu Items */
        div[data-baseweb="popover"],
        div[data-baseweb="popover"] > div,
        ul[data-baseweb="menu"],
        li[data-baseweb="menu-item"] {{
            background-color: #ffffff !important;
            color: #0f172a !important;
        }}

        li[data-baseweb="menu-item"]:hover,
        li[data-baseweb="menu-item"][aria-selected="true"] {{
            background-color: #f1f5f9 !important;
            color: #009bbd !important;
        }}

        /* Inline Code & Badges: soft light blue, never terminal green */
        code {{
            background-color: #f0f9ff !important;
            color: #0284c7 !important;
            border: 1px solid #bae6fd !important;
            border-radius: 6px !important;
            padding: 0.18rem 0.5rem !important;
            font-weight: 600 !important;
            font-size: 0.9rem !important;
            font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace !important;
        }}

        .cycle-badge {{
            display: inline-flex;
            align-items: center;
            background: #f0f9ff;
            border: 1px solid #bae6fd;
            border-radius: 6px;
            padding: 0.35rem 0.85rem;
            font-size: 0.92rem;
            font-weight: 600;
            color: #0284c7;
            letter-spacing: 0.01em;
        }}

        [data-testid="stDataFrame"],
        [data-testid="stTable"] {{
            border: 1px solid #e2e8f0 !important;
            border-radius: 8px !important;
            background: #ffffff !important;
        }}

        /* Streamlit Alerts */
        div[data-testid="stAlert"] {{
            border-radius: 8px !important;
            border: 1px solid #e2e8f0 !important;
            background-color: #ffffff !important;
        }}

        /* Completely remove links, anchor icons, and pointer cursor from all headings */
        [data-testid="stHeaderActionElements"],
        [data-testid="stHeadingWithActionElements"] [data-testid="stHeaderActionElements"],
        [data-testid="stHeadingWithActionElements"] a,
        h1 a, h2 a, h3 a, h4 a, h5 a, h6 a,
        .stHeadingWithActionElements a,
        a.anchorjs-link,
        a.anchor-link {{
            display: none !important;
            visibility: hidden !important;
            pointer-events: none !important;
            opacity: 0 !important;
            width: 0 !important;
            height: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
        }}

        h1, h2, h3, h4, h5, h6,
        [data-testid="stHeadingWithActionElements"],
        [data-testid="stHeadingWithActionElements"] > * {{
            color: #0f172a !important;
            font-weight: 700 !important;
            cursor: default !important;
            user-select: text !important;
        }}

        [data-testid="stHeadingWithActionElements"]:hover [data-testid="stHeaderActionElements"],
        h1:hover a, h2:hover a, h3:hover a, h4:hover a, h5:hover a, h6:hover a {{
            display: none !important;
            visibility: hidden !important;
            opacity: 0 !important;
            pointer-events: none !important;
        }}

        p, span, label {{
            color: #334155 !important;
        }}

        .user-badge {{
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.35rem 0.75rem;
            background: #f0f9ff;
            border: 1px solid #bae6fd;
            border-radius: 6px;
            font-size: 0.85rem;
            font-weight: 600;
            color: #009bbd;
        }}
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


def download_report_excel_api(token: str, job_uuid: str):
    """Download the generated Excel file bytes for a report."""
    try:
        r = requests.get(
            f"{API}/reports/{job_uuid}/download",
            headers={"Authorization": f"Bearer {token}"},
            timeout=45
        )
        if r.ok:
            return r.content
    except Exception:
        pass
    return None


def fetch_dashboard_analytics_api(token: str, financial_year: str = None, month: str = None, job_uuid: str = None):
    """Fetch aggregated dashboard analytics filtered by financial year and month from backend API."""
    try:
        params = {}
        if financial_year and financial_year != "All Financial Years":
            params["financial_year"] = financial_year
        if month and month != "All Months (Full FY)":
            params["month"] = month
        if job_uuid and job_uuid != "All Batches":
            params["job_uuid"] = job_uuid
        r = requests.get(
            f"{API}/analytics/dashboard",
            headers={"Authorization": f"Bearer {token}"},
            params=params,
            timeout=25
        )
        if r.ok:
            return r.json()
    except Exception:
        pass
    return None


def fetch_employees(token: str, search: str = None):
    """Fetch registered employees from the backend API."""
    try:
        params = {}
        if search:
            params["search"] = search
        r = requests.get(f"{API}/employees", headers={"Authorization": f"Bearer {token}"}, params=params, timeout=10)
        if r.ok:
            return r.json()
        if r.status_code != 200:
            err_msg = r.json().get("detail", r.text) if r.headers.get("content-type", "").startswith("application/json") else r.text
            st.error(f"Failed to load employees: {err_msg}")
    except Exception as e:
        st.error(f"Error connecting to backend: {e}")
    return []


def create_employee_api(token: str, employee_code: str, employee_name: str, area: str = None, attendance_mode: str = None):
    """Register a new employee."""
    try:
        payload = {
            "employee_code": employee_code,
            "employee_name": employee_name,
            "area": area if area else None,
            "attendance_mode": attendance_mode if attendance_mode else None
        }
        r = requests.post(f"{API}/employees", headers={"Authorization": f"Bearer {token}"}, json=payload, timeout=10)
        if r.status_code == 201:
            return True, r.json()
        detail = r.json().get("detail", f"Error {r.status_code}") if r.headers.get("content-type", "").startswith("application/json") else r.text
        return False, detail
    except requests.RequestException as e:
        return False, f"Network error: {e}"


def update_employee_api(token: str, employee_id: int, employee_code: str = None, employee_name: str = None, area: str = None, attendance_mode: str = None, is_active: bool = None):
    """Update existing employee details."""
    try:
        payload = {}
        if employee_code is not None:
            payload["employee_code"] = employee_code
        if employee_name is not None:
            payload["employee_name"] = employee_name
        if area is not None:
            payload["area"] = area
        if attendance_mode is not None:
            payload["attendance_mode"] = attendance_mode
        if is_active is not None:
            payload["is_active"] = is_active
        r = requests.put(f"{API}/employees/{employee_id}", headers={"Authorization": f"Bearer {token}"}, json=payload, timeout=10)
        if r.status_code == 200:
            return True, r.json()
        detail = r.json().get("detail", f"Error {r.status_code}") if r.headers.get("content-type", "").startswith("application/json") else r.text
        return False, detail
    except requests.RequestException as e:
        return False, f"Network error: {e}"


def toggle_employee_status_api(token: str, employee_id: int):
    """Toggle employee between Active and Disabled status."""
    try:
        r = requests.post(f"{API}/employees/{employee_id}/toggle-status", headers={"Authorization": f"Bearer {token}"}, timeout=10)
        if r.status_code == 200:
            return True, r.json().get("message", "Status updated successfully.")
        detail = r.json().get("detail", f"Error {r.status_code}") if r.headers.get("content-type", "").startswith("application/json") else r.text
        return False, detail
    except requests.RequestException as e:
        return False, f"Network error: {e}"


def delete_employee_api(token: str, employee_id: int):
    """Disable an employee record by ID."""
    return toggle_employee_status_api(token, employee_id)


def manual_punch_in_api(token: str, employee_code: str, punch_date, punch_time: str):
    """Submit manual punch-in to the backend API."""
    try:
        payload = {
            "employeeCode": employee_code,
            "date": str(punch_date),
            "punchInTime": punch_time
        }
        r = requests.post(
            f"{API}/attendance/manual-punch-in",
            headers={"Authorization": f"Bearer {token}"},
            json=payload,
            timeout=12
        )
        if r.status_code == 200:
            return True, r.json()
        detail = r.json().get("detail", f"Error {r.status_code}") if r.headers.get("content-type", "").startswith("application/json") else r.text
        return False, detail
    except requests.RequestException as e:
        return False, f"Network error: {e}"


def check_existing_punch_api(token: str, employee_code: str, punch_date):
    """Check if attendance already exists for employee and date."""
    try:
        r = requests.get(
            f"{API}/attendance/check",
            headers={"Authorization": f"Bearer {token}"},
            params={"employee_code": employee_code, "date": str(punch_date)},
            timeout=10
        )
        if r.status_code == 200:
            return True, r.json()
        return False, None
    except Exception:
        return False, None


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


def _toggle_sidebar_cb():
    st.session_state.sidebar_open = not st.session_state.get("sidebar_open", True)

def _set_active_tab_cb(tab_name):
    st.session_state.active_tab = tab_name


def render_sidebar():
    """Render modern, collapsible two-mode (expanded & mini icon-dock) sidebar navigation matching user reference screenshot."""
    is_open = st.session_state.get("sidebar_open", True)

    with st.sidebar:
        if not is_open:
            # ── Mini Dock Mode (Collapsed ~68px) ──
            st.button("☰", key="mini_toggle_btn", help="Expand Navigation Menu", on_click=_toggle_sidebar_cb, use_container_width=True)

            st.markdown('<div style="height: 0.4rem; border-bottom: 1px solid #e2e8f0; margin-bottom: 0.5rem;"></div>', unsafe_allow_html=True)

            dash_active = (st.session_state.active_tab == "dashboard")
            emp_active = (st.session_state.active_tab == "employees")
            upload_active = (st.session_state.active_tab == "upload")
            manual_active = (st.session_state.active_tab == "manual_punch")

            user_role = (st.session_state.get("role") or "").lower().replace("_", "-")
            can_manual_punch = user_role in ("coordinator", "dean-faculty")

            st.button("⊞", key="mini_nav_dash", type="primary" if dash_active else "secondary", help="Dashboard", on_click=_set_active_tab_cb, args=("dashboard",), use_container_width=True)
            st.button("👥", key="mini_nav_emp", type="primary" if emp_active else "secondary", help="Employee Management", on_click=_set_active_tab_cb, args=("employees",), use_container_width=True)
            st.button("⤒", key="mini_nav_upload", type="primary" if upload_active else "secondary", help="Upload & Process Attendance", on_click=_set_active_tab_cb, args=("upload",), use_container_width=True)
            if can_manual_punch:
                st.button("⏱️", key="mini_nav_manual", type="primary" if manual_active else "secondary", help="Manual Punch-in", on_click=_set_active_tab_cb, args=("manual_punch",), use_container_width=True)

            return

        # ── Full Expanded Mode (~250px) ──
        logo_data = get_asset_base64("IIITB_logo1.png") or get_asset_base64("iiit_logo.jpg")
        logo_html = (
            f'<img src="{logo_data}" alt="IIIT-B Logo" style="height: 48px; max-width: 100%; object-fit: contain;" />'
            if logo_data else '<span style="font-size: 1.5rem; font-weight: 700; color: #1e3a8a;">IIIT-B</span>'
        )

        col_toggle, col_logo = st.columns([1, 2.5], vertical_alignment="center")
        with col_toggle:
            st.button("☰", key="sidebar_close_btn", help="Collapse Sidebar", on_click=_toggle_sidebar_cb)
        with col_logo:
            st.markdown(f'<div style="display: flex; align-items: center; justify-content: flex-start; padding-left: 0.2rem;">{logo_html}</div>', unsafe_allow_html=True)

        st.markdown(clean_html("""
            <div style="text-align: center; margin: 1.15rem 0 0.85rem 0; font-size: 0.84rem; font-weight: 600; color: #1e3a8a; line-height: 1.4; font-family: 'Inter', system-ui, -apple-system, sans-serif;">
                International Institute of<br/>Information Technology<br/>Bangalore
            </div>
            <div style="border-bottom: 1px solid #e2e8f0; margin-bottom: 0.85rem;"></div>
        """), unsafe_allow_html=True)

        dash_active = (st.session_state.active_tab == "dashboard")
        emp_active = (st.session_state.active_tab == "employees")
        upload_active = (st.session_state.active_tab == "upload")
        manual_active = (st.session_state.active_tab == "manual_punch")

        user_role = (st.session_state.get("role") or "").lower().replace("_", "-")
        can_manual_punch = user_role in ("coordinator", "dean-faculty")

        st.button("⊞   Dashboard", key="side_btn_dashboard", type="primary" if dash_active else "secondary", on_click=_set_active_tab_cb, args=("dashboard",), use_container_width=True)
        st.button("👥   Employees", key="side_btn_employees", type="primary" if emp_active else "secondary", on_click=_set_active_tab_cb, args=("employees",), use_container_width=True)
        st.button("⤒   Upload & Process", key="side_btn_upload", type="primary" if upload_active else "secondary", on_click=_set_active_tab_cb, args=("upload",), use_container_width=True)
        if can_manual_punch:
            st.button("⏱️   Manual Punch-in", key="side_btn_manual", type="primary" if manual_active else "secondary", on_click=_set_active_tab_cb, args=("manual_punch",), use_container_width=True)


def create_plotly_bar_chart(x_vals, y_vals, title, x_title, y_title, is_percentage=False):
    """Create a high-end, responsive Plotly light-themed bar chart."""
    fig = go.Figure()

    hover_template = "%{x}: <b>%{y:.1f}%</b><extra></extra>" if is_percentage else "%{x}: <b>%{y:,}</b><extra></extra>"
    text_labels = [f"{v:.1f}%" if is_percentage else f"{v:,}" for v in y_vals]

    fig.add_trace(go.Bar(
        x=x_vals,
        y=y_vals,
        marker=dict(
            color=y_vals,
            colorscale=[[0, "#00b4d8"], [0.5, "#009bbd"], [1.0, "#0077b6"]],
            line=dict(color="rgba(0, 155, 189, 0.2)", width=1)
        ),
        hovertemplate=hover_template,
        text=text_labels,
        textposition="outside",
        textfont=dict(color="#1e293b", size=11, family="Inter, sans-serif")
    ))

    fig.update_layout(
        title=dict(text=title, font=dict(size=14, color="#0f172a", family="Inter, sans-serif")),
        xaxis=dict(title=x_title, tickfont=dict(color="#475569"), gridcolor="rgba(0,0,0,0.05)"),
        yaxis=dict(title=y_title, tickfont=dict(color="#475569"), gridcolor="rgba(0,0,0,0.06)"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#ffffff",
        margin=dict(l=20, r=20, t=40, b=30),
        height=340,
        showlegend=False
    )
    return fig

def create_plotly_area_chart(x_vals, y_vals, title, x_title, y_title):
    """Create a high-end Plotly light-themed area line trend chart."""
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=x_vals,
        y=y_vals,
        mode="lines+markers",
        fill="tozeroy",
        fillcolor="rgba(0, 155, 189, 0.12)",
        line=dict(color="#009bbd", width=2.5, shape="spline"),
        marker=dict(size=6, color="#0077b6", line=dict(color="#ffffff", width=1.5)),
        hovertemplate="%{x}: <b>%{y:,} Swipes</b><extra></extra>"
    ))

    fig.update_layout(
        title=dict(text=title, font=dict(size=14, color="#0f172a", family="Inter, sans-serif")),
        xaxis=dict(title=x_title, tickfont=dict(color="#475569"), gridcolor="rgba(0,0,0,0.05)"),
        yaxis=dict(title=y_title, tickfont=dict(color="#475569"), gridcolor="rgba(0,0,0,0.06)"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#ffffff",
        margin=dict(l=20, r=20, t=40, b=30),
        height=340
    )
    return fig

def create_plotly_donut_chart(labels, values, title):
    """Create a high-end Plotly light-themed donut chart."""
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.55,
        marker=dict(colors=["#009bbd", "#0284c7", "#38bdf8", "#10b981", "#f59e0b", "#ec4899"]),
        textinfo="percent+label",
        hoverinfo="label+value+percent",
        textfont=dict(color="#ffffff")
    )])

    fig.update_layout(
        title=dict(text=title, font=dict(size=14, color="#0f172a", family="Inter, sans-serif")),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=40, b=20),
        height=340,
        showlegend=True,
        legend=dict(font=dict(color="#475569"))
    )
    return fig

def render_top_bar(title: str, subtitle: str = "", show_refresh: bool = False, refresh_key: str = None) -> bool:
    """Render top teal app bar matching the reference screenshot with user avatar and Logout button."""
    user = st.session_state.get("user", "coordinator")
    role = st.session_state.get("role", "coordinator")
    initial = (user[0] if user else "C").upper()

    # Clean title without leading emojis for the top teal bar
    clean_title = re.sub(r'^[^\w\s]+\s*', '', title).strip()
    key_slug = re.sub(r'[^a-zA-Z0-9]', '_', clean_title.lower())[:12]

    # Top Teal App Bar Banner with C badge and clean Logout button
    st.markdown(clean_html(f"""
        <div class="top-teal-bar">
            <div class="top-teal-title">{clean_title}</div>
            <div class="top-teal-right">
                <span class="user-name-label">{user}</span>
                <span class="user-avatar-badge" title="{user} ({role})">{initial}</span>
                <a href="?action=logout" class="top-teal-logout-btn" title="Sign out" target="_self">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: -2px; margin-right: 4px;"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path><polyline points="16 17 21 12 16 7"></polyline><line x1="21" y1="12" x2="9" y2="12"></line></svg>Logout
                </a>
            </div>
        </div>
    """), unsafe_allow_html=True)

    refreshed = False
    if show_refresh:
        col_sub, col_ref = st.columns([9.5, 1.5])
        with col_sub:
            if subtitle:
                st.markdown(f"<div style='font-size: 0.88rem; color: #64748b; margin: 0.35rem 0 1rem 0.2rem;'>{subtitle}</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div style='margin-bottom: 0.4rem;'></div>", unsafe_allow_html=True)
        with col_ref:
            st.markdown("<div style='padding-top: 0.1rem;'></div>", unsafe_allow_html=True)
            if st.button("↻ Refresh", key=refresh_key or f"refr_{key_slug}", use_container_width=True):
                refreshed = True
    else:
        if subtitle:
            st.markdown(f"<div style='font-size: 0.88rem; color: #64748b; margin: 0.35rem 0 1rem 0.2rem;'>{subtitle}</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div style='margin-bottom: 0.4rem;'></div>", unsafe_allow_html=True)

    return refreshed


def render_dashboard():
    refreshed = render_top_bar(
        title="Attendance Analytics Dashboard",
        subtitle="Analyze employee attendance trends financial year-wise and month-wise across institutional cycles.",
        show_refresh=True,
        refresh_key="refresh_dash_btn"
    )
    if refreshed:
        st.session_state.pop("dash_analytics_data", None)
        st.rerun()

    # Read active filters from session state if previously set
    selected_fy = st.session_state.get("dash_selected_fy")
    selected_month = st.session_state.get("dash_selected_month")

    # Fetch analytics
    data = fetch_dashboard_analytics_api(
        st.session_state.token,
        financial_year=selected_fy,
        month=selected_month
    )

    if not data or data.get("total_dates", 0) == 0:
        st.markdown(clean_html("""
            <div class="glass-card" style="text-align: center; padding: 3rem 2rem; margin-top: 1rem;">
                <h3 style="color: #475569; margin-bottom: 0.5rem; font-size: 1.25rem;">No Attendance Data Found</h3>
                <p style="color: #64748b; font-size: 0.95rem; max-width: 480px; margin: 0 auto;">
                    No attendance records match the selected financial year or period.
                    Upload Attendance &amp; Swipe Excel files or add manual punches to generate analytics.
                </p>
            </div>
        """), unsafe_allow_html=True)

        c1, c2, c3, c4 = st.columns(4)
        for col, val, label in [
            (c1, "—", "Total Employees"), (c2, "—", "Dates Analyzed"),
            (c3, "—%", "Punch-In Rate"), (c4, "Active", "System Status")
        ]:
            with col:
                st.markdown(f'<div class="stat-box"><div class="stat-val">{val}</div><div class="stat-label">{label}</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Go to Upload & Process →", type="primary"):
            st.session_state.active_tab = "upload"
            st.rerun()
        return

    # ─── Financial Year & Month Filter Controls ──────────────────────────────
    fys = data.get("financial_years", ["FY 2026-27"])
    active_fy = data.get("selected_financial_year", fys[0] if fys else "FY 2026-27")
    fy_idx = fys.index(active_fy) if active_fy in fys else 0

    months_list = ["All Months (Full FY)"] + (data.get("available_months") or [])
    active_month = data.get("selected_month", "All Months (Full FY)")
    month_idx = months_list.index(active_month) if active_month in months_list else 0

    st.markdown(clean_html("""
        <div style="margin-top: 0.25rem; margin-bottom: 0.5rem;">
            <span style="font-weight: 700; font-size: 1.05rem; color: #0f172a; font-family: 'Inter', sans-serif;">
                🎯 Analytics Filter
            </span>
            <span style="font-size: 0.85rem; color: #64748b; margin-left: 0.5rem;">
                Select Financial Year and drill down into specific months
            </span>
        </div>
    """), unsafe_allow_html=True)

    fcol1, fcol2 = st.columns([1, 1])
    with fcol1:
        new_fy = st.selectbox(
            "📅 Financial Year:",
            options=fys,
            index=fy_idx,
            key="dash_selected_fy"
        )
    with fcol2:
        new_month = st.selectbox(
            "🗓️ Month Filter:",
            options=months_list,
            index=month_idx,
            key="dash_selected_month"
        )

    # If the user changed the selectbox during interaction, reload if needed
    if new_fy != active_fy or new_month != active_month:
        updated_data = fetch_dashboard_analytics_api(
            st.session_state.token,
            financial_year=new_fy,
            month=new_month
        )
        if updated_data:
            data = updated_data
            active_fy = data.get("selected_financial_year", new_fy)
            active_month = data.get("selected_month", new_month)

    # ─── Scope Header Banner ────────────────────────────────────────────────
    st.markdown(clean_html(f"""
        <div class="glass-card" style="margin: 0.85rem 0 1.25rem 0; padding: 0.85rem 1.25rem; border-left: 4px solid #009bbd; display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 0.5rem;">
            <div style="display: flex; align-items: center; gap: 0.65rem; flex-wrap: wrap;">
                <span style="font-size: 1.1rem; font-weight: 700; color: #0f172a;">
                    📅 {active_fy}
                </span>
                <span style="color: #cbd5e1;">|</span>
                <span style="font-size: 0.95rem; color: #334155; font-weight: 600;">
                    🗓️ {active_month}
                </span>
            </div>
            <div style="font-size: 0.84rem; color: #0369a1; font-weight: 600; background: #e0f2fe; padding: 0.3rem 0.75rem; border-radius: 9999px;">
                {data.get('total_dates', 0)} Active Dates &nbsp;•&nbsp; {data.get('total_employees', 0)} Employees
            </div>
        </div>
    """), unsafe_allow_html=True)

    # ─── 4 Key Metric Cards ──────────────────────────────────────────────────
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(f'<div class="stat-box"><div class="stat-val">{data["total_employees"]}</div><div class="stat-label">Total Employees</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="stat-box"><div class="stat-val">{data["total_dates"]}</div><div class="stat-label">Dates Analyzed</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="stat-box"><div class="stat-val">{data["total_punches"]:,} / {data["total_possible"]:,}</div><div class="stat-label">Matched Punch-Ins</div></div>', unsafe_allow_html=True)
    with m4:
        st.markdown(f'<div class="stat-box"><div class="stat-val">{data["completion_rate"]}%</div><div class="stat-label">Punch-In Completion Rate</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='margin-bottom: 1.25rem;'></div>", unsafe_allow_html=True)

    # ─── MONTHWISE ANALYTICS ─────────────────────────────────────────────────
    st.markdown(f"### 🗓️ Financial Year Month-Wise Analytics ({active_fy})")
    month_summary = data.get("month_summary", [])
    if month_summary:
        col_m1, col_m2 = st.columns([1, 1])
        with col_m1:
            st.markdown("#### Monthly Summary")
            month_table_data = []
            for m_info in month_summary:
                month_table_data.append({
                    "Month": m_info["month_name"],
                    "Active Dates": m_info["dates_count"],
                    "Punch-Ins Recorded": f"{m_info['punches']:,}",
                    "Expected Slots": f"{m_info['total']:,}",
                    "Completion Rate": f"{m_info['completion_rate']}%"
                })
            st.dataframe(month_table_data, use_container_width=True)
        with col_m2:
            m_names = [m["month_short"] + " (" + m["month_name"].split()[-1] + ")" if " " in m["month_name"] else m["month_name"] for m in month_summary]
            m_rates = [m["completion_rate"] for m in month_summary]
            fig_month = create_plotly_bar_chart(
                m_names,
                m_rates,
                f"Month-Wise Completion Rate (%) — {active_fy}",
                "Month",
                "Completion Rate (%)",
                is_percentage=True
            )
            st.plotly_chart(fig_month, use_container_width=True)
    else:
        st.info("No monthly breakdown available for this selection.")

    st.markdown("<div style='margin-bottom: 1.25rem;'></div>", unsafe_allow_html=True)

    # ─── DATEWISE ANALYTICS ──────────────────────────────────────────────────
    view_title = active_month if active_month != "All Months (Full FY)" else f"All Dates in {active_fy}"
    st.markdown(f"### 📈 Date-Wise Daily Trends ({view_title})")
    daily_counts = data.get("daily_counts", {})
    if daily_counts:
        date_labels = [d.get("date_display", k) for k, d in daily_counts.items()]
        daily_punches = [d["punches"] for d in daily_counts.values()]
        fig_date = create_plotly_area_chart(
            date_labels,
            daily_punches,
            f"Daily Punch-In Volume Trend ({view_title})",
            "Date",
            "Punch-Ins"
        )
        st.plotly_chart(fig_date, use_container_width=True)

    st.markdown("#### Daily Attendance Breakdown")
    daily_table = []
    for k, d_info in daily_counts.items():
        punches = d_info["punches"]
        tot = d_info["total"]
        pct = round((punches / tot * 100), 1) if tot > 0 else 0
        first_punch = d_info.get("first_punch") or "—"
        last_punch = d_info.get("last_punch") or "—"
        daily_table.append({
            "Date": d_info.get("date_display", k),
            "Day": d_info.get("day_name", "—"),
            "Month": d_info.get("month_name", "—"),
            "Punch-Ins": f"{punches:,}",
            "Total Employees": f"{tot:,}",
            "Attendance %": f"{pct}%",
            "Earliest Punch-In": first_punch,
            "Latest Punch-In": last_punch
        })
    st.dataframe(daily_table, use_container_width=True)

    # Per-employee per-date detail table
    emp_daily_records = data.get("employee_daily_records", [])
    if emp_daily_records:
        st.markdown("<div style='margin-top: 1rem;'></div>", unsafe_allow_html=True)
        col_hdr, col_srch = st.columns([3, 2])
        with col_hdr:
            st.markdown("#### Detailed Punch-In Roster")
        with col_srch:
            search_emp = st.text_input(
                "Search employee",
                placeholder="🔍 Search employee name or code...",
                label_visibility="collapsed",
                key=f"dash_emp_search_{active_fy}_{active_month}"
            )

        if search_emp:
            se = search_emp.strip().lower()
            emp_daily_records = [
                r for r in emp_daily_records
                if se in str(r.get("employee_code", "")).lower() or se in str(r.get("employee_name", "")).lower()
            ]

        emp_detail_table = []
        for row in emp_daily_records:
            emp_detail_table.append({
                "Date": row.get("date_display") or row.get("date", "—"),
                "Employee Code": row.get("employee_code", ""),
                "Employee Name": row.get("employee_name") or "—",
                "Punch-In Time": row.get("punch_in_time", "—"),
                "Status": row.get("status", "—")
            })
        st.dataframe(emp_detail_table, use_container_width=True, height=450)

    st.markdown("<div style='margin-bottom: 1.25rem;'></div>", unsafe_allow_html=True)

    # ─── EMPLOYEE INSIGHTS ───────────────────────────────────────────────────
    st.markdown(f"### 👥 Employee Attendance Insights ({active_month})")
    ec1, ec2 = st.columns([1, 1])
    with ec1:
        st.markdown("#### 🏆 Highest Attendance Frequency")
        top_data = data.get("top_performers", [])
        if top_data:
            st.dataframe(top_data, use_container_width=True)
        else:
            st.info("No employee punch-in data available.")

    with ec2:
        st.markdown("#### ⚠️ Employees with Missing / Incomplete Punch-Ins")
        missing_data = data.get("missing_punch_employees", [])
        if missing_data:
            st.dataframe(missing_data, use_container_width=True)
        else:
            st.success("🎉 All employees have complete punch-in records!")


def render_upload():
    render_top_bar(
        title="Upload & Process Attendance",
        subtitle="Upload Attendance Master file and Swipe file to automatically insert Punchin Times.",
        show_refresh=False
    )

    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(clean_html("""
            <div class="glass-card">
                <h4 style="color: #0f172a; font-size: 1.05rem; margin: 0; font-weight: 600;">1. Attendance Master Excel</h4>
            </div>
        """), unsafe_allow_html=True)
        af = st.file_uploader("Select Attendance File", type=["xlsx", "xlsm"], key="att_file")
        ap = st.text_input("Attendance Password (if encrypted)", type="password", placeholder="Leave blank if not password protected", key="att_pass")

    with col2:
        st.markdown(clean_html("""
            <div class="glass-card">
                <h4 style="color: #0f172a; font-size: 1.05rem; margin: 0; font-weight: 600;">2. Swipe Log Excel</h4>
            </div>
        """), unsafe_allow_html=True)
        sf = st.file_uploader("Select Swipe Log File", type=["xlsx", "xlsm"], key="swp_file")
        sp = st.text_input("Swipe Log Password (if encrypted)", type="password", placeholder="Leave blank if not password protected", key="swp_pass")

    st.markdown("<br>", unsafe_allow_html=True)
    
    st.info("**Rule**: Employee Code + Date are matched across files. The earliest swipe timestamp is picked as Punchin Time. Dates without swipes remain blank.")

    if st.button("Generate & Process Punch-In Excel", type="primary", use_container_width=True):
        if not af or not sf:
            st.error("Please select both the Attendance Excel file and Swipe Log Excel file.")
            return
        
        with st.spinner("Processing Excel spreadsheets, parsing swipes, and calculating punch-in times..."):
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
                    
                    st.success("Updated Attendance Excel generated successfully!")

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

                    col_dl, col_nav1 = st.columns([2, 1])
                    with col_dl:
                        st.download_button(
                            label="Download Updated Attendance Excel",
                            data=BytesIO(r.content),
                            file_name="Updated_Attendance_Punchin.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            type="primary"
                        )
                    with col_nav1:
                        if st.button("View Dashboard →", use_container_width=True, key="post_proc_dash_btn"):
                            st.session_state.active_tab = "dashboard"
                            st.rerun()
                elif r.status_code == 401:
                    st.session_state.clear()
                    st.error("Session expired. Please sign in again.")
                    st.rerun()
                else:
                    detail = r.json().get("detail", "Processing failed.")
                    st.error(f"Error: {detail}")
            except requests.RequestException as e:
                st.error(f"Backend connection error: {e}")

    # ── Render Batch Execution Logs & Generated Reports directly below upload ──
    st.markdown("<div style='border-bottom: 2px solid #e2e8f0; margin: 2.2rem 0 1.5rem 0;'></div>", unsafe_allow_html=True)
    render_batch_reports_section(in_upload_page=True)


def render_batch_reports_section(in_upload_page: bool = True):
    """Render the Batch Execution Logs & Generated Reports table and interactive on-screen Excel viewer."""
    kp = "up" if in_upload_page else "rep"

    st.markdown(clean_html("""
        <div style="margin-top: 0.5rem; margin-bottom: 0.85rem;">
            <h3 style="color: #0f172a; font-size: 1.35rem; font-weight: 700; margin: 0; font-family: 'Inter', sans-serif;">
                📋 Batch Execution Logs &amp; Generated Reports
            </h3>
            <p style="color: #64748b; font-size: 0.88rem; margin-top: 0.25rem;">
                Review past processing batches, inspect the full generated Excel spreadsheet on-screen, and download the processed workbooks.
            </p>
        </div>
    """), unsafe_allow_html=True)

    # Load available reports from backend database
    if st.session_state.get("available_reports") is None:
        with st.spinner("Loading report history from database..."):
            st.session_state.available_reports = fetch_reports(st.session_state.token)

    db_reports = st.session_state.get("available_reports") or []

    # Map database records into unified display records
    unified_records = []
    seen_uuids = set()
    for r in db_reports:
        job_uuid = r.get("job_uuid", "")
        raw_created = r.get("created_at") or ""
        created_display = raw_created[:19].replace("T", " ") if raw_created else "—"
        if job_uuid:
            seen_uuids.add(job_uuid)
        unified_records.append({
            "job_uuid": job_uuid,
            "created_at": created_display,
            "cycle_label": r.get("cycle_label", "—"),
            "attendance_file": r.get("attendance_filename") or "—",
            "swipe_file": r.get("swipe_filename") or "—",
            "valid_swipes": r.get("valid_swipes_count", 0),
            "invalid_swipes": r.get("invalid_swipes_count", 0),
            "date_count": r.get("date_count", 0),
            "status": r.get("status", "COMPLETED"),
            "username": r.get("username", "coordinator")
        })

    # Merge session history
    for h in st.session_state.get("history", []):
        ts = h.get("timestamp", "")
        matching = any(u["created_at"].startswith(ts[:16]) for u in unified_records)
        if not matching:
            unified_records.append({
                "job_uuid": "",
                "created_at": ts,
                "cycle_label": "Recent Batch",
                "attendance_file": h.get("attendance_file", "—"),
                "swipe_file": h.get("swipe_file", "—"),
                "valid_swipes": h.get("valid_swipes", 0),
                "invalid_swipes": h.get("invalid_swipes", 0),
                "date_count": h.get("date_count", 0),
                "status": "COMPLETED",
                "username": st.session_state.get("user", "coordinator")
            })

    if not unified_records:
        st.markdown(clean_html("""
            <div class="glass-card" style="text-align: center; padding: 2.5rem 1.5rem; margin-top: 0.75rem;">
                <h4 style="color: #475569; margin-bottom: 0.35rem; font-size: 1.15rem;">No Processing Batches Found</h4>
                <p style="color: #64748b; font-size: 0.9rem; max-width: 480px; margin: 0 auto;">
                    No report runs have been recorded in the database yet.
                    Upload Attendance &amp; Swipe Excel files above to generate your first report.
                </p>
            </div>
        """), unsafe_allow_html=True)
        return

    # Aggregate Analytics KPIs
    total_runs = len(unified_records)
    total_valid = sum(h.get("valid_swipes", 0) for h in unified_records)
    total_invalid = sum(h.get("invalid_swipes", 0) for h in unified_records)
    total_dates = sum(h.get("date_count", 0) for h in unified_records)

    rc1, rc2, rc3, rc4 = st.columns(4)
    with rc1:
        st.metric("Total Batches Processed", total_runs)
    with rc2:
        st.metric("Cumulative Valid Swipes", f"{total_valid:,}")
    with rc3:
        st.metric("Cumulative Invalid Swipes", f"{total_invalid:,}")
    with rc4:
        st.metric("Total Dates Updated", f"{total_dates:,}")

    st.markdown("<div style='margin-bottom: 0.75rem;'></div>", unsafe_allow_html=True)

    # Display Batch Execution Logs table
    display_df = [{
        "Execution Time": r["created_at"],
        "Cycle": r["cycle_label"],
        "Attendance File": r["attendance_file"],
        "Swipe File": r["swipe_file"],
        "Valid Swipes": f"{r['valid_swipes']:,}",
        "Invalid / Skipped": f"{r['invalid_swipes']:,}",
        "Dates Updated": r["date_count"],
        "Processed By": r["username"],
        "Status": r["status"]
    } for r in unified_records]

    st.dataframe(display_df, use_container_width=True)

    # ── Report Inspector: Click/Select Report to View Whole Generated Excel ──
    db_options = [r for r in unified_records if r.get("job_uuid")]
    if not db_options:
        return

    st.markdown("<div style='margin-top: 1.5rem; margin-bottom: 0.5rem;'></div>", unsafe_allow_html=True)

    # Report selection options
    report_map = {}
    for r in db_options:
        label = f"{r['cycle_label']} — {r['attendance_file']} ({r['created_at']})"
        report_map[label] = r["job_uuid"]

    selected_report_label = st.selectbox(
        "Select Report Batch to View & Download Generated Excel:",
        options=list(report_map.keys()),
        key=f"report_selector_{kp}"
    )

    selected_uuid = report_map.get(selected_report_label)
    if not selected_uuid:
        return

    # Fetch report detail
    detail_cache_key = f"detail_{selected_uuid}"
    if detail_cache_key not in st.session_state or st.session_state[detail_cache_key] is None:
        with st.spinner("Loading generated Excel spreadsheet..."):
            st.session_state[detail_cache_key] = fetch_report_detail(selected_uuid, st.session_state.token)

    detail = st.session_state.get(detail_cache_key)
    if not detail:
        st.error("Failed to load details for the selected report.")
        return

    # Metadata Card
    st.markdown(clean_html(f"""
        <div class="glass-card" style="margin: 0.75rem 0 1rem 0; padding: 1rem 1.25rem; border-left: 4px solid #009bbd;">
            <div style="display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 0.5rem;">
                <div>
                    <div style="font-size: 1.1rem; font-weight: 700; color: #0f172a;">
                        📄 Generated Attendance Workbook: {detail.get('attendance_filename', 'Attendance File')}
                    </div>
                    <div style="color: #64748b; font-size: 0.88rem; margin-top: 0.15rem;">
                        Cycle: <strong>{detail.get('cycle_label', '—')}</strong> &nbsp;|&nbsp; 
                        Employees: <strong>{detail.get('total_employees', 0):,}</strong> &nbsp;|&nbsp; 
                        Dates: <strong>{detail.get('total_dates', 0)}</strong> &nbsp;|&nbsp; 
                        Valid Punches: <strong>{detail.get('total_punches', 0):,} / {detail.get('total_possible', 0):,} ({detail.get('completion_rate', 0)}%)</strong>
                    </div>
                </div>
            </div>
        </div>
    """), unsafe_allow_html=True)

    # Action Toolbar: Download button + Search Box
    col_dl_act, col_search = st.columns([3, 5])
    with col_dl_act:
        excel_bytes = download_report_excel_api(st.session_state.token, selected_uuid)
        if excel_bytes:
            filename = f"Updated_Attendance_{detail.get('attendance_filename', selected_uuid[:8]).replace('.xlsx', '').replace('.xlsm', '')}.xlsx"
            st.download_button(
                label="📥 Download Generated Excel (.xlsx)",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                key=f"dl_excel_btn_{selected_uuid}_{kp}"
            )
        else:
            st.button("📥 Download Generated Excel (.xlsx)", disabled=True, use_container_width=True, key=f"dl_excel_dis_{selected_uuid}_{kp}")

    with col_search:
        search_filter = st.text_input(
            "Search in Sheet",
            placeholder="🔍 Search employee name or code in sheet...",
            label_visibility="collapsed",
            key=f"search_sheet_input_{selected_uuid}_{kp}"
        )

    # Display Mode Toggle
    view_mode = st.radio(
        "Display Mode:",
        ["📊 Full Excel Grid (Date Columns)", "📋 Chronological Punch Records"],
        horizontal=True,
        key=f"view_mode_radio_{selected_uuid}_{kp}"
    )

    if view_mode == "📊 Full Excel Grid (Date Columns)":
        grid_rows = detail.get("excel_grid_rows") or []
        if search_filter:
            sf_lower = search_filter.strip().lower()
            grid_rows = [
                r for r in grid_rows
                if sf_lower in str(r.get("Employee Code", "")).lower() or sf_lower in str(r.get("Employee Name", "")).lower()
            ]

        if grid_rows:
            st.dataframe(grid_rows, use_container_width=True, height=520)
            st.caption(f"Showing {len(grid_rows)} employees across {len(detail.get('excel_date_columns', []))} date columns. You can scroll horizontally to view all daily punch-in times.")
        else:
            st.info("No matching employees found in this report.")

    else:
        # Chronological list view
        daily_records = detail.get("employee_daily_records") or []
        if search_filter:
            sf_lower = search_filter.strip().lower()
            daily_records = [
                r for r in daily_records
                if sf_lower in str(r.get("employee_code", "")).lower() or sf_lower in str(r.get("employee_name", "")).lower()
            ]

        if daily_records:
            st.dataframe(daily_records, use_container_width=True, height=520)
            st.caption(f"Showing {len(daily_records):,} daily attendance rows.")
        else:
            st.info("No matching records found in this report.")


def render_reports():
    st.session_state.active_tab = "upload"
    render_upload()


def render_employees():
    refreshed = render_top_bar(
        title="Employee Directory & Management",
        subtitle="Manage employee profiles, register personnel, and maintain attendance master details.",
        show_refresh=True,
        refresh_key="refresh_emp_btn"
    )

    if refreshed or "employee_list" not in st.session_state or st.session_state.employee_list is None:
        with st.spinner("Loading employee directory..."):
            st.session_state.employee_list = fetch_employees(st.session_state.token)

    employees = st.session_state.get("employee_list") or []

    # KPI Statistics Row
    total_emp = len(employees)
    active_count = sum(1 for e in employees if e.get("is_active", True) is not False)
    disabled_count = total_emp - active_count
    unique_areas = sorted(list({e.get("area") for e in employees if e.get("area") and str(e.get("area")).strip()}))
    unique_modes = sorted(list({e.get("attendance_mode") for e in employees if e.get("attendance_mode") and str(e.get("attendance_mode")).strip()}))

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("Total Employees", f"{total_emp:,}")
    with m2:
        st.metric("Active Personnel", f"{active_count:,}")
    with m3:
        st.metric("Disabled Personnel", f"{disabled_count:,}")
    with m4:
        st.metric("Departments / Areas", len(unique_areas))

    st.markdown("<div style='margin-bottom: 1rem;'></div>", unsafe_allow_html=True)

    # ── State Initializations ──
    if "emp_show_add" not in st.session_state:
        st.session_state.emp_show_add = False
    if "emp_editing_id" not in st.session_state:
        st.session_state.emp_editing_id = None
    if "emp_disabling_id" not in st.session_state:
        st.session_state.emp_disabling_id = None
    if "emp_deleting_id" in st.session_state and st.session_state.emp_deleting_id is not None:
        st.session_state.emp_disabling_id = st.session_state.emp_deleting_id
        st.session_state.emp_deleting_id = None

    # ── Disable Confirmation Card ──
    if st.session_state.emp_disabling_id is not None:
        dis_target = next((e for e in employees if e["id"] == st.session_state.emp_disabling_id), None)
        if dis_target:
            st.markdown(clean_html(f"""
                <div class="glass-card" style="border-left: 4px solid #f59e0b; background: #fffbeb !important;">
                    <div style="font-weight: 700; color: #b45309; font-size: 1.05rem; margin-bottom: 0.35rem;">
                        ⚠️ Confirm Employee Deactivation
                    </div>
                    <div style="color: #475569; font-size: 0.95rem; margin-bottom: 0.85rem;">
                        Are you sure you want to disable <strong>{dis_target['employee_name']}</strong> 
                        (Employee Code: <code>{dis_target['employee_code']}</code>)? The employee will be marked as disabled and can be re-enabled at any time.
                    </div>
                </div>
            """), unsafe_allow_html=True)
            col_d1, col_d2, _ = st.columns([2, 2, 6])
            with col_d1:
                if st.button("🚫 Yes, Disable", type="primary", key="confirm_disable_btn", use_container_width=True):
                    ok, msg = toggle_employee_status_api(st.session_state.token, dis_target["id"])
                    if ok:
                        st.session_state.emp_disabling_id = None
                        st.session_state.employee_list = None
                        st.toast(f"Employee {dis_target['employee_name']} disabled.", icon="🚫")
                        st.rerun()
                    else:
                        st.error(msg)
            with col_d2:
                if st.button("Cancel", key="cancel_disable_btn", use_container_width=True):
                    st.session_state.emp_disabling_id = None
                    st.rerun()
            st.markdown("<div style='margin-bottom: 1rem;'></div>", unsafe_allow_html=True)

    # ── Edit Employee Form ──
    if st.session_state.emp_editing_id is not None:
        edit_target = next((e for e in employees if e["id"] == st.session_state.emp_editing_id), None)
        if edit_target:
            st.markdown(clean_html(f"""
                <div class="glass-card" style="border-left: 4px solid #009bbd; margin-bottom: 1rem;">
                    <div style="font-weight: 700; color: #0f172a; font-size: 1.1rem; margin-bottom: 0.2rem;">
                        ✏️ Edit Employee Profile
                    </div>
                    <div style="color: #64748b; font-size: 0.88rem;">
                        Updating details for {edit_target['employee_name']} (ID: {edit_target['id']})
                    </div>
                </div>
            """), unsafe_allow_html=True)

            with st.form("edit_employee_form"):
                col_e1, col_e2 = st.columns(2)
                with col_e1:
                    edit_name = st.text_input("Employee Name *", value=edit_target["employee_name"], key="edit_name_inp")
                    edit_code = st.text_input("Employee Code *", value=edit_target["employee_code"], key="edit_code_inp")
                with col_e2:
                    edit_area = st.text_input("Department / Area", value=edit_target.get("area") or "", placeholder="e.g. Computer Science, Administration", key="edit_area_inp")
                    mode_options = ["Biometric", "RFID Card", "Face Recognition", "Mobile App", "Manual"]
                    current_mode = edit_target.get("attendance_mode") or "Biometric"
                    curr_idx = mode_options.index(current_mode) if current_mode in mode_options else 0
                    edit_mode = st.selectbox("Attendance Mode", mode_options, index=curr_idx, key="edit_mode_inp")
                    target_active = edit_target.get("is_active", True) is not False
                    edit_status = st.selectbox("Account Status", ["Active", "Disabled"], index=0 if target_active else 1, key="edit_status_inp")

                col_b1, col_b2, _ = st.columns([2, 2, 6])
                with col_b1:
                    save_edit_submitted = st.form_submit_button("💾 Save Changes", type="primary", use_container_width=True)
                with col_b2:
                    cancel_edit = st.form_submit_button("Cancel", use_container_width=True)

                if save_edit_submitted:
                    if not edit_name.strip():
                        st.error("Employee Name cannot be empty.")
                    elif not edit_code.strip():
                        st.error("Employee Code cannot be empty.")
                    else:
                        ok, res = update_employee_api(
                            st.session_state.token,
                            edit_target["id"],
                            employee_code=edit_code.strip(),
                            employee_name=edit_name.strip(),
                            area=edit_area.strip() if edit_area.strip() else None,
                            attendance_mode=edit_mode,
                            is_active=(edit_status == "Active")
                        )
                        if ok:
                            st.session_state.emp_editing_id = None
                            st.session_state.employee_list = None
                            st.toast("Employee details updated successfully!", icon="✅")
                            st.rerun()
                        else:
                            st.error(res)

                if cancel_edit:
                    st.session_state.emp_editing_id = None
                    st.rerun()

            st.markdown("<div style='margin-bottom: 1.5rem;'></div>", unsafe_allow_html=True)

    # ── Add Employee Panel ──
    if st.session_state.emp_show_add:
        st.markdown(clean_html("""
            <div class="glass-card" style="border-left: 4px solid #10b981; margin-bottom: 1rem;">
                <div style="font-weight: 700; color: #0f172a; font-size: 1.1rem; margin-bottom: 0.2rem;">
                    ➕ Register New Employee
                </div>
                <div style="color: #64748b; font-size: 0.88rem;">
                    Enter the employee's name and unique code to register them into the database.
                </div>
            </div>
        """), unsafe_allow_html=True)

        with st.form("add_employee_form", clear_on_submit=True):
            col_a1, col_a2 = st.columns(2)
            with col_a1:
                new_name = st.text_input("Employee Name *", placeholder="e.g. Dr. Ramesh Kumar", key="add_name_inp")
                new_code = st.text_input("Employee Code *", placeholder="e.g. EMP1042 or 1042", key="add_code_inp")
            with col_a2:
                new_area = st.text_input("Department / Area", placeholder="e.g. Computer Science, Administration", key="add_area_inp")
                new_mode = st.selectbox("Attendance Mode", ["Biometric", "RFID Card", "Face Recognition", "Mobile App", "Manual"], index=0, key="add_mode_inp")

            col_sub1, col_sub2, _ = st.columns([2, 2, 6])
            with col_sub1:
                add_submitted = st.form_submit_button("➕ Save Employee", type="primary", use_container_width=True)
            with col_sub2:
                add_cancelled = st.form_submit_button("Cancel", use_container_width=True)

            if add_submitted:
                if not new_name.strip():
                    st.error("Please provide Employee Name.")
                elif not new_code.strip():
                    st.error("Please provide Employee Code.")
                else:
                    ok, res = create_employee_api(
                        st.session_state.token,
                        employee_code=new_code.strip(),
                        employee_name=new_name.strip(),
                        area=new_area.strip() if new_area.strip() else None,
                        attendance_mode=new_mode
                    )
                    if ok:
                        st.session_state.emp_show_add = False
                        st.session_state.employee_list = None
                        st.toast(f"Employee {new_name.strip()} added successfully!", icon="✅")
                        st.rerun()
                    else:
                        st.error(res)

            if add_cancelled:
                st.session_state.emp_show_add = False
                st.rerun()

        st.markdown("<div style='margin-bottom: 1.5rem;'></div>", unsafe_allow_html=True)

    # ── Search & Actions Bar ──
    col_act, col_search, col_filter, col_status = st.columns([2.2, 3.8, 2.3, 1.7])
    with col_act:
        add_btn_label = "✖ Close Form" if st.session_state.emp_show_add else "➕ Add New Employee"
        if st.button(add_btn_label, type="primary" if not st.session_state.emp_show_add else "secondary", key="toggle_add_btn", use_container_width=True):
            st.session_state.emp_show_add = not st.session_state.emp_show_add
            st.rerun()

    with col_search:
        search_query = st.text_input("🔍 Search employees", placeholder="Search by name or code...", label_visibility="collapsed", key="emp_search_box")

    with col_filter:
        filter_options = ["All Departments"] + unique_areas
        selected_dept = st.selectbox("Department", filter_options, label_visibility="collapsed", key="emp_dept_filter")

    with col_status:
        selected_status = st.selectbox("Status", ["All Status", "Active", "Disabled"], label_visibility="collapsed", key="emp_status_filter")

    # ── Filter Logic ──
    filtered_emps = employees
    if search_query:
        q = search_query.strip().lower()
        filtered_emps = [e for e in filtered_emps if q in e.get("employee_name", "").lower() or q in e.get("employee_code", "").lower()]
    if selected_dept != "All Departments":
        filtered_emps = [e for e in filtered_emps if e.get("area") == selected_dept]
    if selected_status == "Active":
        filtered_emps = [e for e in filtered_emps if e.get("is_active", True) is not False]
    elif selected_status == "Disabled":
        filtered_emps = [e for e in filtered_emps if e.get("is_active", True) is False]

    st.markdown("<div style='margin-top: 0.5rem;'></div>", unsafe_allow_html=True)

    # ── Directory Table & Operations ──
    if not filtered_emps:
        st.markdown(clean_html("""
            <div class="glass-card" style="text-align: center; padding: 3rem 2rem; margin-top: 1rem;">
                <div style="font-size: 2.5rem; margin-bottom: 0.5rem;">👥</div>
                <h3 style="color: #475569; margin-bottom: 0.5rem; font-size: 1.25rem;">No Employees Found</h3>
                <p style="color: #64748b; font-size: 0.95rem; max-width: 480px; margin: 0 auto;">
                    No employee records matched your search or department filter. Click <strong>Add New Employee</strong> to register personnel.
                </p>
            </div>
        """), unsafe_allow_html=True)
        return

    # Count display & Export action
    col_hdr, col_exp = st.columns([7, 3])
    with col_hdr:
        st.markdown(f"<div style='font-size: 0.95rem; font-weight: 600; color: #475569; padding-top: 0.5rem;'>Showing <strong>{len(filtered_emps)}</strong> of <strong>{total_emp}</strong> registered employees</div>", unsafe_allow_html=True)
    with col_exp:
        # Generate CSV download
        import io, csv
        csv_buffer = io.StringIO()
        writer = csv.writer(csv_buffer)
        writer.writerow(["ID", "Employee Code", "Employee Name", "Department/Area", "Attendance Mode", "Status", "Created At"])
        for e in filtered_emps:
            e_status = "Active" if e.get("is_active", True) is not False else "Disabled"
            writer.writerow([e.get("id"), e.get("employee_code"), e.get("employee_name"), e.get("area") or "", e.get("attendance_mode") or "", e_status, e.get("created_at") or ""])
        st.download_button(
            label="📥 Export Directory (CSV)",
            data=csv_buffer.getvalue(),
            file_name="employees_directory.csv",
            mime="text/csv",
            use_container_width=True,
            key="export_emp_csv"
        )

    st.markdown("<div style='margin-bottom: 0.5rem;'></div>", unsafe_allow_html=True)

    # Table Header Row
    th_col1, th_col2, th_col3, th_col4, th_col5, th_col6 = st.columns([1.5, 3.5, 2.5, 2, 1.2, 1.2])
    with th_col1:
        st.markdown("<div style='font-size: 0.78rem; font-weight: 700; color: #64748b; text-transform: uppercase;'>Code</div>", unsafe_allow_html=True)
    with th_col2:
        st.markdown("<div style='font-size: 0.78rem; font-weight: 700; color: #64748b; text-transform: uppercase;'>Name</div>", unsafe_allow_html=True)
    with th_col3:
        st.markdown("<div style='font-size: 0.78rem; font-weight: 700; color: #64748b; text-transform: uppercase;'>Department / Area</div>", unsafe_allow_html=True)
    with th_col4:
        st.markdown("<div style='font-size: 0.78rem; font-weight: 700; color: #64748b; text-transform: uppercase;'>Mode</div>", unsafe_allow_html=True)
    with th_col5:
        st.markdown("<div style='font-size: 0.78rem; font-weight: 700; color: #64748b; text-transform: uppercase; text-align: center;'>Edit</div>", unsafe_allow_html=True)
    with th_col6:
        st.markdown("<div style='font-size: 0.78rem; font-weight: 700; color: #64748b; text-transform: uppercase; text-align: center;'>Disable</div>", unsafe_allow_html=True)

    st.markdown("<div style='border-bottom: 2px solid #e2e8f0; margin-bottom: 0.6rem;'></div>", unsafe_allow_html=True)

    # Pagination
    page_size = 50
    total_pages = max(1, (len(filtered_emps) + page_size - 1) // page_size)
    page = 1
    if total_pages > 1:
        page = st.number_input(f"Page (1 to {total_pages})", min_value=1, max_value=total_pages, value=1, step=1, key="emp_page_num")

    start_idx = (page - 1) * page_size
    end_idx = start_idx + page_size
    page_emps = filtered_emps[start_idx:end_idx]

    # Render each row
    for emp in page_emps:
        r_col1, r_col2, r_col3, r_col4, r_col5, r_col6 = st.columns([1.5, 3.5, 2.5, 2, 1.2, 1.2])
        
        emp_code = emp.get("employee_code", "—")
        emp_name = emp.get("employee_name", "—")
        emp_area = emp.get("area") or "—"
        emp_mode = emp.get("attendance_mode") or "Biometric"
        emp_id = emp["id"]
        is_active = emp.get("is_active", True) is not False

        with r_col1:
            st.markdown(f"<div style='padding-top: 0.35rem;'><code>{emp_code}</code></div>", unsafe_allow_html=True)
        with r_col2:
            if is_active:
                st.markdown(f"<div style='padding-top: 0.35rem; font-weight: 600; color: #0f172a;'>{emp_name}</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div style='padding-top: 0.35rem; font-weight: 600; color: #94a3b8;'>{emp_name} <span style='display:inline-block; font-size:0.72rem; padding:1px 6px; border-radius:4px; background:#fee2e2; color:#b91c1c; margin-left:4px; font-weight:700;'>DISABLED</span></div>", unsafe_allow_html=True)
        with r_col3:
            st.markdown(f"<div style='padding-top: 0.35rem; color: #475569;'>{emp_area}</div>", unsafe_allow_html=True)
        with r_col4:
            st.markdown(f"<div style='padding-top: 0.35rem;'><span class='cycle-badge' style='padding: 0.15rem 0.5rem; font-size: 0.8rem;'>{emp_mode}</span></div>", unsafe_allow_html=True)
        with r_col5:
            if st.button("✏️", key=f"btn_edit_{emp_id}", help=f"Edit {emp_name}", use_container_width=True):
                st.session_state.emp_editing_id = emp_id
                st.session_state.emp_show_add = False
                st.session_state.emp_disabling_id = None
                st.rerun()
        with r_col6:
            if is_active:
                if st.button("🚫", key=f"btn_dis_{emp_id}", help=f"Disable {emp_name}", use_container_width=True):
                    st.session_state.emp_disabling_id = emp_id
                    st.session_state.emp_editing_id = None
                    st.session_state.emp_show_add = False
                    st.rerun()
            else:
                if st.button("✅", key=f"btn_ena_{emp_id}", help=f"Enable {emp_name}", use_container_width=True):
                    ok, msg = toggle_employee_status_api(st.session_state.token, emp_id)
                    if ok:
                        st.session_state.employee_list = None
                        st.toast(f"Employee {emp_name} re-enabled!", icon="✅")
                        st.rerun()
                    else:
                        st.error(msg)

        st.markdown("<div style='border-bottom: 1px solid #f1f5f9; margin-bottom: 0.35rem;'></div>", unsafe_allow_html=True)


def render_manual_punch():
    """Render the Manual Punch-in page matching the application UI design and role constraints."""
    refreshed = render_top_bar(
        title="⏱️ Manual Punch-in",
        subtitle="Manually record or update employee attendance punch-in times.",
        show_refresh=True,
        refresh_key="refresh_manual_punch_btn"
    )

    user_role = (st.session_state.get("role") or "").lower().replace("_", "-")
    if user_role not in ("coordinator", "dean-faculty"):
        st.markdown(clean_html("""
            <div class="glass-card" style="border-left: 4px solid #ef4444; background: #fff5f5 !important; margin-top: 1rem;">
                <div style="font-weight: 700; color: #b91c1c; font-size: 1.05rem; margin-bottom: 0.35rem;">
                    🚫 Access Denied
                </div>
                <div style="color: #475569; font-size: 0.95rem;">
                    Manual punch-in is strictly restricted to <strong>Coordinator</strong> and <strong>Dean-Faculty</strong> roles.
                </div>
            </div>
        """), unsafe_allow_html=True)
        return

    # Load employees
    if refreshed or "employee_list" not in st.session_state or st.session_state.employee_list is None:
        with st.spinner("Loading employee list..."):
            st.session_state.employee_list = fetch_employees(st.session_state.token)

    employees = st.session_state.get("employee_list") or []

    if not employees:
        st.markdown(clean_html("""
            <div class="glass-card" style="border-left: 4px solid #f59e0b; background: #fffbeb !important; margin-top: 1rem;">
                <div style="font-weight: 700; color: #b45309; font-size: 1.05rem; margin-bottom: 0.35rem;">
                    ⚠️ No Employees Available
                </div>
                <div style="color: #475569; font-size: 0.95rem;">
                    No registered employees found in the master records. Please register employees in the <strong>Employees</strong> section or upload an attendance sheet first.
                </div>
            </div>
        """), unsafe_allow_html=True)
        return

    # Prepare searchable employee options: Employee Name (Code) - Area (active only)
    emp_options = {}
    emp_display_list = []
    for emp in employees:
        if emp.get("is_active", True) is False:
            continue
        code = emp.get("employee_code", "")
        name = emp.get("employee_name", "")
        area = emp.get("area", "")
        display = f"{name} ({code})" + (f" — {area}" if area else "")
        emp_options[display] = emp
        emp_display_list.append(display)

    # Clean Card Header
    st.markdown(clean_html("""
        <div class="glass-card" style="margin-bottom: 1.25rem;">
            <div style="font-weight: 700; color: #0f172a; font-size: 1.15rem; margin-bottom: 0.25rem;">
                ✍️ Record Attendance Punch-in
            </div>
            <div style="color: #64748b; font-size: 0.88rem;">
                Select an employee, punch-in date, and punch-in time. If a punch-in record already exists for the selected date, it will be safely updated without creating duplicate entries.
            </div>
        </div>
    """), unsafe_allow_html=True)

    # Input Fields in 3 responsive columns
    col_emp, col_date, col_time = st.columns([4.5, 3, 2.5])

    with col_emp:
        st.markdown("<label style='font-size: 0.9rem; font-weight: 600; color: #1e293b; display: block; margin-bottom: 6px;'>Employee *</label>", unsafe_allow_html=True)
        selected_display = st.selectbox(
            "Employee",
            options=emp_display_list,
            index=0,
            label_visibility="collapsed",
            key="manual_punch_emp_select"
        )
        selected_emp = emp_options.get(selected_display) if selected_display else None
        selected_code = selected_emp["employee_code"] if selected_emp else None

    with col_date:
        st.markdown("<label style='font-size: 0.9rem; font-weight: 600; color: #1e293b; display: block; margin-bottom: 6px;'>Punch-in Date *</label>", unsafe_allow_html=True)
        selected_date = st.date_input(
            "Punch-in Date",
            value=date.today(),
            label_visibility="collapsed",
            key="manual_punch_date_picker"
        )

    now_clean = datetime.now().time().replace(second=0, microsecond=0)
    with col_time:
        st.markdown("<label style='font-size: 0.9rem; font-weight: 600; color: #1e293b; display: block; margin-bottom: 6px;'>Punch-in Time *</label>", unsafe_allow_html=True)
        selected_time = st.time_input(
            "Punch-in Time",
            value=now_clean,
            step=60,
            label_visibility="collapsed",
            key="manual_punch_time_picker"
        )

    selected_time_clean = selected_time.replace(second=0, microsecond=0) if selected_time else now_clean
    formatted_time_str = selected_time_clean.strftime("%H:%M:00")
    selected_hm = selected_time_clean.strftime("%H:%M")

    # Real-time duplicate attendance check & indicator
    if selected_code and selected_date:
        has_record, check_data = check_existing_punch_api(st.session_state.token, selected_code, selected_date)
        if has_record and check_data and check_data.get("exists"):
            fp = check_data.get("first_punch") or "—"
            st_val = check_data.get("status") or "ABSENT"
            existing_hm = fp[:5] if len(fp) >= 5 else fp

            # Check if user's selected punch-in time matches existing punch-in time
            is_same_punch_time = bool(existing_hm and selected_hm == existing_hm)

            if is_same_punch_time:
                # Punch-in time is not changed
                st.markdown(clean_html(f"""
                    <div class="glass-card" style="border-left: 4px solid #f59e0b; background: #fffbeb !important; margin: 0.85rem 0; padding: 0.85rem 1.15rem;">
                        <div style="font-weight: 600; color: #b45309; font-size: 0.92rem; margin-bottom: 0.25rem;">
                            ℹ️ Punch-in Time Unchanged
                        </div>
                        <div style="color: #475569; font-size: 0.88rem; line-height: 1.5;">
                            Employee <strong>{selected_emp['employee_name']}</strong> (<code>{selected_code}</code>) already has a punch-in recorded at <strong>{fp}</strong> on <strong>{selected_date.strftime('%d-%b-%Y')}</strong>:
                            <span style="display: inline-block; margin-left: 6px; padding: 2px 8px; border-radius: 4px; background: #fde68a; color: #92400e; font-weight: 600;">Status: {st_val}</span>
                            <br/>The selected time ({selected_hm}) is already recorded for this date.
                        </div>
                    </div>
                """), unsafe_allow_html=True)
            else:
                # User selected a DIFFERENT punch-in time: show clean update notice without alarming duplicate warning
                st.markdown(clean_html(f"""
                    <div class="glass-card" style="border-left: 4px solid #009bbd; background: #f0f9ff !important; margin: 0.85rem 0; padding: 0.85rem 1.15rem;">
                        <div style="font-weight: 600; color: #0369a1; font-size: 0.92rem; margin-bottom: 0.25rem;">
                            ✏️ Update Punch-in Time
                        </div>
                        <div style="color: #475569; font-size: 0.88rem; line-height: 1.5;">
                            Employee <strong>{selected_emp['employee_name']}</strong> (<code>{selected_code}</code>) has existing punch: <strong>{fp}</strong> on {selected_date.strftime('%d-%b-%Y')}.
                            <br/>Submitting will update the punch-in time to <strong>{selected_hm}:00</strong>.
                        </div>
                    </div>
                """), unsafe_allow_html=True)
        else:
            st.markdown(clean_html(f"""
                <div class="glass-card" style="border-left: 4px solid #10b981; background: #f0fdf4 !important; margin: 0.85rem 0; padding: 0.85rem 1.15rem;">
                    <div style="font-weight: 600; color: #047857; font-size: 0.92rem; margin-bottom: 0.2rem;">
                        ✨ Ready to Punch In
                    </div>
                    <div style="color: #475569; font-size: 0.88rem;">
                        No prior punch-in found for <strong>{selected_emp['employee_name']}</strong> on {selected_date.strftime('%d-%b-%Y')}.
                        Submitting will create a new attendance record marked <strong>PRESENT</strong> at <strong>{selected_hm}:00</strong>.
                    </div>
                </div>
            """), unsafe_allow_html=True)

    st.markdown("<div style='margin-bottom: 0.85rem;'></div>", unsafe_allow_html=True)

    # Submit Button
    col_btn, col_rest = st.columns([3, 7])
    with col_btn:
        submit_clicked = st.button(
            "Add Punch-in",
            type="primary",
            use_container_width=True,
            key="btn_add_manual_punch"
        )

    if submit_clicked:
        if not selected_code:
            st.error("Please select a valid employee.")
            return
        if not selected_date:
            st.error("Please select a valid punch-in date.")
            return
        if not selected_time:
            st.error("Please select a valid punch-in time.")
            return

        with st.spinner("Submitting manual punch-in..."):
            ok, result = manual_punch_in_api(
                token=st.session_state.token,
                employee_code=selected_code,
                punch_date=selected_date,
                punch_time=formatted_time_str
            )

        if ok:
            msg = result.get("message", "Manual punch-in added successfully")
            rec = result.get("record") or {}
            action_type = result.get("action", "created")

            st.toast(msg, icon="✅")
            st.markdown(clean_html(f"""
                <div class="glass-card" style="border-left: 4px solid #10b981; background: #f0fdf4 !important; margin-top: 1rem;">
                    <div style="font-weight: 700; color: #047857; font-size: 1.1rem; margin-bottom: 0.4rem;">
                        ✅ {msg}
                    </div>
                    <div style="color: #334155; font-size: 0.92rem; line-height: 1.7;">
                        <strong>Employee:</strong> {rec.get('employee_name', selected_emp['employee_name'])} (<code>{rec.get('employee_code', selected_code)}</code>)<br/>
                        <strong>Date:</strong> {rec.get('date', str(selected_date))}<br/>
                        <strong>Punch-in Time:</strong> <code style="font-weight: 700; color: #0284c7; font-size: 0.95rem;">{rec.get('punch_in_time', formatted_time_str)}</code><br/>
                        <strong>Status:</strong> <span style="padding: 2px 8px; border-radius: 4px; background: #dcfce7; color: #166534; font-weight: 600;">{rec.get('status', 'PRESENT')}</span><br/>
                        <strong>Operation:</strong> {action_type.capitalize()}
                    </div>
                </div>
            """), unsafe_allow_html=True)
        else:
            st.error(f"Failed to record punch-in: {result}")


def handle_logout_query_param():
    """Handle instant logout triggered from the top navbar link."""
    try:
        if st.query_params.get("action") in ["logout", ["logout"]]:
            st.query_params.clear()
            st.session_state.clear()
            st.rerun()
    except AttributeError:
        try:
            if st.experimental_get_query_params().get("action") in ["logout", ["logout"]]:
                st.experimental_set_query_params()
                st.session_state.clear()
                st.rerun()
        except Exception:
            pass


def main():
    handle_logout_query_param()
    inject_custom_css(login_page=not bool(st.session_state.token))

    if not st.session_state.token:
        render_login()
        return

    # Always render sidebar (it renders full 250px or mini 68px icon dock depending on session state)
    render_sidebar()

    # Render selected page tab
    if st.session_state.active_tab == "dashboard":
        render_dashboard()
    elif st.session_state.active_tab == "employees":
        render_employees()
    elif st.session_state.active_tab == "upload":
        render_upload()
    elif st.session_state.active_tab in ("report", "reports"):
        st.session_state.active_tab = "upload"
        render_upload()
    elif st.session_state.active_tab == "manual_punch":
        render_manual_punch()

if __name__ == "__main__":
    main()