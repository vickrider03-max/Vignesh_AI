# ==============================
# APP ENTRY / AUTH / GLOBAL CONTROL
# Modular entry point generated from legacy_app.py.
# app.py keeps global Streamlit orchestration, authentication, logout, preview
# routing, sidebar upload control, active-tab routing, and shared initialization.
# Heavy backend logic lives in functions.py; tab UI lives in tab_*.py.
# ==============================

import functions as fn
from functions import *
from router import TAB_OPTIONS, init_router, render_tab_router
from state_firewall import init_state_firewall, tab_state_scope
from tab_memory import init_tab_memory
from tab_chat import render_chat_tab
from tab_dashboard import render_dashboard_tab
from tab_compare import render_compare_tab
from tab_capl import render_capl_tab

# -------------------------------
# IMPORTS
# -------------------------------
# Standard library and third-party imports for the application.



# -------------------------------
# GLOBAL VARIABLES & CONSTANTS
# -------------------------------
# Persistent storage for document previews across app reruns.
# fn.PREVIEW_TOKENS maps tokens to metadata, fn.PREVIEW_STORE holds file data.
fn.PREVIEW_TOKENS = {}  # token -> {'file_name': str, 'timestamp': datetime}
fn.PREVIEW_STORE = {}   # token -> file_dict

APP_DIR = os.path.dirname(os.path.abspath(__file__))
PREVIEW_DATA_FILE = os.path.join(APP_DIR, "preview_data.pkl")
WORKSPACE_DB_FILE = os.path.join(APP_DIR, "workspace_memory.db")
WORKSPACE_MEMORY_KEY = "workspace_memory"
PDF_PREVIEW_RESOLUTION = 100
PDF_PREVIEW_WINDOW = 25
PDF_ASSET_SCAN_PAGE_LIMIT = 10
MAX_VECTOR_TEXT_CHARS = 250000

# ===============================================
# CACHING SYSTEM FOR PERFORMANCE OPTIMIZATION
# ===============================================
# This caching layer dramatically improves load times by avoiding redundant processing


# Global cache instances
fn.FILE_TEXT_CACHE = CacheManager(max_size=100)
fn.VECTOR_STORE_CACHE = CacheManager(max_size=20)
fn.EXCEL_DATA_CACHE = CacheManager(max_size=50)
fn.EMBEDDINGS_CACHE = CacheManager(max_size=200)

# Cache for file hashes to detect modifications
fn.FILE_HASH_CACHE = {}

# ===============================================
# HASH-BASED CHANGE DETECTION
# ===============================================


# ===============================================
# PREVIEW PERSISTENCE HELPERS
# -------------------------------
# Functions to save/load preview data so document previews persist across Streamlit reruns.
# This allows users to share preview links that work even after app restart.





# -------------------------------
# MERCEDES LOGO GENERATION
# -------------------------------
# Generates an animated GIF of the Mercedes-Benz logo using matplotlib.
# Used in the sidebar and login screen for branding.


# -------------------------------
# STREAMLIT PAGE CONFIG
# -------------------------------
st.set_page_config(page_title="🧠 IntelliDoc AI ", layout="wide")

# Mobile viewport meta tag for proper scaling
st.markdown(
    """
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    """,
    unsafe_allow_html=True,
)

# -------------------------------
# GLOBAL CSS STYLING
# -------------------------------
# Custom CSS to hide Streamlit branding, style buttons, and add animations.
# Applied globally to override default Streamlit appearance.
st.markdown(
    """
    <style>
        #MainMenu,
        header,
        footer,
        [data-testid="stToolbar"],
        [data-testid="stFooter"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        .viewerBadge_container__1QSob,
        .css-1lsmgbg,
        .css-rtg1gx,
        .stAppDeployButton {
            display: none !important;
            visibility: hidden !important;
            opacity: 0 !important;
            height: 0 !important;
            width: 0 !important;
            pointer-events: none !important;
        }
    </style>
    <script>
        const hideStreamlitBranding = () => {
            const selectors = [
                '#MainMenu',
                'header',
                'footer',
                '[data-testid="stToolbar"]',
                '[data-testid="stFooter"]',
                '[data-testid="stDecoration"]',
                '[data-testid="stStatusWidget"]',
                '.viewerBadge_container__1QSob',
                '.css-1lsmgbg',
                '.css-rtg1gx',
                '.stAppDeployButton'
            ];

            selectors.forEach(selector => {
                document.querySelectorAll(selector).forEach(el => {
                    el.style.setProperty('display', 'none', 'important');
                    el.style.setProperty('visibility', 'hidden', 'important');
                    el.style.setProperty('opacity', '0', 'important');
                    el.style.setProperty('height', '0', 'important');
                    el.style.setProperty('width', '0', 'important');
                    el.style.setProperty('pointer-events', 'none', 'important');
                });
            });

            document.querySelectorAll('*').forEach(el => {
                try {
                    const text = (el.innerText || el.textContent || '').trim();
                    if (/made with streamlit/i.test(text) || /github/i.test(text)) {
                        el.style.setProperty('display', 'none', 'important');
                        el.style.setProperty('visibility', 'hidden', 'important');
                        el.style.setProperty('opacity', '0', 'important');
                        el.style.setProperty('height', '0', 'important');
                        el.style.setProperty('width', '0', 'important');
                        el.style.setProperty('pointer-events', 'none', 'important');
                    }
                } catch (err) {
                    // ignore inaccessible nodes
                }
            });

            document.querySelectorAll('img, svg').forEach(el => {
                try {
                    const src = (el.src || '') + (el.outerHTML || '');
                    if (/streamlit/i.test(src) || /github/i.test(src)) {
                        el.style.setProperty('display', 'none', 'important');
                        el.style.setProperty('visibility', 'hidden', 'important');
                        el.style.setProperty('opacity', '0', 'important');
                        el.style.setProperty('height', '0', 'important');
                        el.style.setProperty('width', '0', 'important');
                        el.style.setProperty('pointer-events', 'none', 'important');
                    }
                } catch (err) {
                    // ignore inaccessible nodes
                }
            });
        };

        const observer = new MutationObserver(hideStreamlitBranding);
        observer.observe(document.documentElement, { childList: true, subtree: true });
        hideStreamlitBranding();
        setInterval(hideStreamlitBranding, 1000);
    </script>
    """,
    unsafe_allow_html=True,
)

# Load preview data from file
load_preview_data()

# Clean up expired preview tokens on app start
cleanup_expired_preview_tokens()

try:
    logo_data = get_needle_minimalist_logo()
except Exception:
    logo_data = None

# -------------------------------
# ADDITIONAL CSS STYLING
# -------------------------------
# More CSS for dashboard grids, metric cards, loading animations, and responsive design.
st.markdown(
    """
    <style>
        :root {
            --primary: #e8f6ff;
            --secondary: #64748b;
            --background: #ffffff;
            --surface: #f8fafc;
            --text: #1e293b;
            --text-secondary: #64748b;
            --border: #e2e8f0;
            --success: #10b981;
            --warning: #f59e0b;
            --error: #ef4444;
            --button-bg: #e8f6ff;
            --button-hover: #d0e8f8;
            --button-text: #1e293b;
        }
        
        *, *::before, *::after {
            box-sizing: border-box;
        }

        html, body {
            min-width: 0;
            overflow-x: hidden;
            overflow-y: auto !important;
            height: auto !important;
            background: var(--background)!important;
            color: var(--text);
            transition: background 0.3s ease, color 0.3s ease;
        }

        div[role="main"], section.main, .stApp {
            min-width: 0;
            max-width: 1600px;
            width: 100%;
            margin: 0 auto;
        }

        .block-container {
            padding-left: clamp(1rem, 2vw, 1.75rem) !important;
            padding-right: clamp(1rem, 2vw, 1.75rem) !important;
            max-width: 1600px;
        }

        .stSidebar {
            min-width: clamp(220px, 18vw, 300px) !important;
            max-width: clamp(260px, 20vw, 380px) !important;
            width: clamp(220px, 18vw, 360px) !important;
        }

        .stSidebarNav {
            min-width: clamp(220px, 18vw, 280px) !important;
            width: 100% !important;
        }

        /* Main content positioning */
        .main .block-container {
            margin-left: 0 !important;
            margin-right: 0 !important;
            padding-left: 1rem !important;
            padding-right: 1rem !important;
            width: 100% !important;
            max-width: none !important;
        }

        /* Ensure main content is visible */
        section.main {
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
        }

        /* Responsive spacing for smaller and larger screens */
        @media (min-width: 640px) {
            .main .block-container {
                padding-left: 1.5rem !important;
                padding-right: 1.5rem !important;
            }
        }

        @media (min-width: 900px) {
            .main .block-container {
                padding-left: 2rem !important;
                padding-right: 2rem !important;
            }
        }

        @media (min-width: 1200px) {
            .main .block-container {
                padding-left: 3rem !important;
                padding-right: 3rem !important;
            }
        }

        @media (max-width: 900px) {
            .stSidebar,
            .stSidebarNav {
                width: 100% !important;
                min-width: 0 !important;
                max-width: 100% !important;
            }
            .stSidebar {
                position: relative !important;
                transform: none !important;
            }
        }

        @media (min-width: 1024px) {
            .main .block-container {
                padding-left: 3rem !important;
                padding-right: 3rem !important;
            }
        }

        /* Streamlit page container fixes */
        .stApp,
        div[data-testid="stAppViewContainer"] {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            min-height: auto !important;
        }

        section.main,
        .main,
        .main .block-container,
        div[data-testid="stAppViewContainer"],
        div[data-testid="stMainContent"],
        div[data-testid="main"] {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
            overflow-x: hidden !important;
            overflow-y: visible !important;
        }

        section.main {
            display: block !important;
            min-height: auto !important;
        }

        .main .block-container {
            max-width: 100% !important;
            width: 100% !important;
        }

        /* Make sure sidebar and content layout stays aligned */
        div[data-testid="stAppViewContainer"] > div {
            min-width: 0 !important;
            overflow-x: hidden !important;
        }

        .dashboard-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 1rem;
            margin: 1rem 0;
        }
        
        .metric-card {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 1.5rem;
            text-align: center;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        
        .metric-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        
        .card-label {
            font-size: 0.875rem;
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 0.5rem;
        }
        
        .card-value {
            font-size: 1.5rem;
            font-weight: 700;
            color: var(--primary);
        }
        
        .spinner {
            width: 40px;
            height: 40px;
            border: 3px solid var(--border);
            border-top: 3px solid var(--primary);
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        }
        
        .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(255, 255, 255, 0.9);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 9999;
            animation: fadeIn 0.3s ease;
        }
        
        .loading-content {
            text-align: center;
            padding: 2rem;
            background: var(--surface);
            border-radius: 12px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.1);
            animation: slideUp 0.4s ease;
        }
        
        .loading-dots {
            display: inline-block;
            width: 20px;
            height: 20px;
            border-radius: 50%;
            background: var(--primary);
            animation: bounce 1.4s ease-in-out infinite both;
            margin: 0 2px;
        }
        
        .loading-dots:nth-child(1) { animation-delay: -0.32s; }
        .loading-dots:nth-child(2) { animation-delay: -0.16s; }
        
        /* Button styling - Ultra-aggressive Streamlit override */
        .stButton > button {
            background-color: #e8f6ff !important;
            color: #1e293b !important;
            border: 2px solid #c0dff0 !important;
            border-radius: 6px !important;
            padding: 0.5rem 1rem !important;
            font-weight: 600 !important;
            transition: all 0.2s ease !important;
            box-shadow: 0 2px 4px rgba(200, 230, 250, 0.2) !important;
        }
        
        .stButton > button:hover {
            background-color: #d0e8f8 !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 4px 8px rgba(175, 215, 245, 0.3) !important;
            border-color: #a0c8e8 !important;
        }
        
        .stButton > button:active {
            transform: translateY(0) !important;
            box-shadow: 0 2px 4px rgba(200, 230, 250, 0.2) !important;
        }
        
        /* Override Streamlit's internal button styling */
        div.stButton > button,
        div[data-testid="stBaseButton"] button,
        button[kind="primary"],
        button[kind="secondary"],
        button[kind="tertiary"] {
            background-color: #e8f6ff !important;
            color: #1e293b !important;
            border: 2px solid #c0dff0 !important;
        }
        
        div.stButton > button:hover,
        div[data-testid="stBaseButton"] button:hover,
        button[kind="primary"]:hover,
        button[kind="secondary"]:hover,
        button[kind="tertiary"]:hover {
            background-color: #d0e8f8 !important;
        }
        
        /* Alternative Streamlit button selectors */
        [role="button"],
        [data-testid*="button"] {
            background-color: #e8f6ff !important;
            color: #1e293b !important;
        }
        
        [role="button"]:hover,
        [data-testid*="button"]:hover {
            background-color: #d0e8f8 !important;
        }
        
        /* Strip Streamlit theme blue and apply light blue */
        button {
            background-color: #e8f6ff !important;
            color: #1e293b !important;
            border-color: #c0dff0 !important;
        }
        
        button:hover {
            background-color: #d0e8f8 !important;
        }
        
        /* Smaller buttons for logout and reset */
        button[data-testid*="main_logout_btn"],
        button[data-testid*="reset_chat_selection"],
        button[data-testid*="reset_dashboard_selection"],
        button[data-testid*="reset_compare_selection"],
        button[data-testid*="reset_capl_selection"] {
            padding: 0.25rem 0.75rem !important;
            font-size: 0.875rem !important;
            min-width: auto !important;
            width: auto !important;
        }
        
        /* Animations */
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        @keyframes slideUp {
            from { 
                opacity: 0; 
                transform: translateY(20px); 
            }
            to { 
                opacity: 1; 
                transform: translateY(0); 
            }
        }
        
        @keyframes bounce {
            0%, 80%, 100% { 
                transform: scale(0);
                opacity: 0.5;
            } 
            40% { 
                transform: scale(1);
                opacity: 1;
            }
        }
        
        @keyframes pulse {
            0% { transform: scale(1); }
            50% { transform: scale(1.05); }
            100% { transform: scale(1); }
        }
        
        @keyframes slideInLeft {
            from { 
                opacity: 0; 
                transform: translateX(-30px); 
            }
            to { 
                opacity: 1; 
                transform: translateX(0); 
            }
        }
        
        @keyframes slideInRight {
            from { 
                opacity: 0; 
                transform: translateX(30px); 
            }
            to { 
                opacity: 1; 
                transform: translateX(0); 
            }
        }
        
        @keyframes glow {
            0% { box-shadow: 0 0 5px rgba(168, 216, 240, 0.3); }
            50% { box-shadow: 0 0 20px rgba(127, 197, 232, 0.6); }
            100% { box-shadow: 0 0 5px rgba(168, 216, 240, 0.3); }
        }
        
        @keyframes wiggle {
            0%, 100% { transform: rotate(0deg); }
            25% { transform: rotate(-3deg); }
            75% { transform: rotate(3deg); }
        }
        
        @keyframes float {
            0%, 100% { transform: translateY(0px); }
            50% { transform: translateY(-10px); }
        }
        
        .fade-in {
            animation: fadeIn 0.5s ease-out;
        }
        
        .slide-in-left {
            animation: slideInLeft 0.6s ease-out;
        }
        
        .slide-in-right {
            animation: slideInRight 0.6s ease-out;
        }
        
        .pulse {
            animation: pulse 2s infinite;
        }
        
        .glow {
            animation: glow 3s infinite;
        }
        
        .wiggle {
            animation: wiggle 1s ease-in-out;
        }
        
        .float {
            animation: float 3s ease-in-out infinite;
        }
        
        @media (max-width: 768px) {
            .dashboard-grid {
                grid-template-columns: 1fr;
            }
        }
        
        /* Hide Streamlit footer and GitHub icon */
        footer,
        body > footer,
        footer *,
        [data-testid="stFooter"],
        [data-testid="stFooter"] *,
        [data-testid="stDecoration"],
        [data-testid="stDecoration"] *,
        [data-testid="stToolbar"],
        [data-testid="stHeader"],
        [data-testid="stStatusWidget"],
        #MainMenu,
        #MainMenu *,
        a[href*="streamlit.io"],
        a[href*="github.com"],
        a[href*="github"],
        [title*="GitHub"],
        [aria-label*="GitHub"],
        [aria-label*="Streamlit"],
        [role="complementary"] img,
        [role="complementary"] svg,
        [style*="position: fixed"][style*="bottom: 0px"][style*="right: 0px"],
        [style*="position: fixed"][style*="bottom: 1px"][style*="right: 1px"],
        [style*="position: fixed"][style*="bottom: 10px"][style*="right: 10px"],
        [style*="position: fixed"][style*="bottom: 8px"][style*="right: 8px"],
        [data-testid="stDecoration"][style*="position: fixed"] {
            display: none !important;
            visibility: hidden !important;
            opacity: 0 !important;
            height: 0 !important;
            width: 0 !important;
            pointer-events: none !important;
        }
    </style>
    <script>
        const hideStreamlitFooter = () => {
            const selectors = [
                'footer',
                'footer *',
                '[data-testid="stFooter"]',
                '[data-testid="stDecoration"]',
                '[data-testid="stToolbar"]',
                '[data-testid="stHeader"]',
                '#MainMenu',
                'a[href*="streamlit.io"]',
                'a[href*="github.com"]',
                'a[href*="github"]',
                '[aria-label*="GitHub"]',
                '[aria-label*="Streamlit"]',
                '[title*="GitHub"]',
                '[title*="Streamlit"]'
            ];
            selectors.forEach(selector => {
                document.querySelectorAll(selector).forEach(el => {
                    el.style.setProperty('display', 'none', 'important');
                    el.style.setProperty('visibility', 'hidden', 'important');
                    el.style.setProperty('opacity', '0', 'important');
                    el.style.setProperty('height', '0', 'important');
                    el.style.setProperty('width', '0', 'important');
                    el.style.setProperty('pointer-events', 'none', 'important');
                });
            });

            document.querySelectorAll('*').forEach(el => {
                try {
                    const text = (el.innerText || el.textContent || '').trim();
                    const hasIcon = el.querySelector('img[src*="streamlit"], img[src*="github"], svg');
                    if (text.includes('Made with Streamlit') || text.includes('GitHub') || hasIcon) {
                        el.style.setProperty('display', 'none', 'important');
                        el.style.setProperty('visibility', 'hidden', 'important');
                        el.style.setProperty('opacity', '0', 'important');
                        el.style.setProperty('height', '0', 'important');
                        el.style.setProperty('width', '0', 'important');
                        el.style.setProperty('pointer-events', 'none', 'important');
                    }
                } catch (err) {
                    // ignore inaccessible nodes
                }
            });

            document.querySelectorAll('body *').forEach(el => {
                try {
                    const style = window.getComputedStyle(el);
                    if (style.position === 'fixed') {
                        const rect = el.getBoundingClientRect();
                        if (rect.bottom >= window.innerHeight - 90 && rect.right >= window.innerWidth - 90 && rect.width <= 90 && rect.height <= 90) {
                            if (el.querySelector('img, svg') || /Streamlit|GitHub/i.test(el.innerText || el.textContent || '')) {
                                el.style.setProperty('display', 'none', 'important');
                                el.style.setProperty('visibility', 'hidden', 'important');
                                el.style.setProperty('opacity', '0', 'important');
                                el.style.setProperty('height', '0', 'important');
                                el.style.setProperty('width', '0', 'important');
                                el.style.setProperty('pointer-events', 'none', 'important');
                            }
                        }
                    }
                } catch (err) {
                    // ignore inaccessible nodes
                }
            });
        };

        const footerObserver = new MutationObserver(() => hideStreamlitFooter());
        footerObserver.observe(document.body, { childList: true, subtree: true });
        hideStreamlitFooter();
        setInterval(hideStreamlitFooter, 1000);

        const applyLightButtonStyles = () => {
            const selectors = [
                'button',
                'input[type="button"]',
                'input[type="submit"]',
                '[role="button"]',
                '.stButton > button',
                'div[data-testid="stBaseButton"] button'
            ];
            document.querySelectorAll(selectors.join(',')).forEach(el => {
                el.style.setProperty('background-color', '#e8f6ff', 'important');
                el.style.setProperty('color', '#1e293b', 'important');
                el.style.setProperty('border-color', '#c0dff0', 'important');
                el.style.setProperty('border-style', 'solid', 'important');
                el.style.setProperty('border-width', '2px', 'important');
                el.style.setProperty('box-shadow', '0 2px 4px rgba(200, 230, 250, 0.2)', 'important');
            });
        };

        const buttonObserver = new MutationObserver(() => {
            applyLightButtonStyles();
        });
        buttonObserver.observe(document.body, { childList: true, subtree: true });
        applyLightButtonStyles();
    </script>
    """,
    unsafe_allow_html=True
)

# Ensure session state keys exist before rendering login status
for key, default_value in [
    ("is_authenticated", False),
    ("logged_in_username", ""),
    ("user_role", None),
    ("login_history", []),
    ("selected_files", []),
    ("file_texts", {}),
    ("vector_stores", {}),
    ("chat_file_selection", []),
    ("chat_summary_downloads", {"images": [], "tables": [], "csv": [], "diagrams": []}),
    ("chat_item_downloads", {"csv": [], "diagrams": []}),
    ("messages", []),
    ("welcome_shown", False),
    ("mobile_sidebar_visible", False),
]:
    if key not in st.session_state:
        st.session_state[key] = default_value













# ============================================
# SIMPLE HEADER - Moved higher for better visibility
# ============================================
if st.session_state.is_authenticated:
    st.markdown(
        """
        <style>
            section.main .block-container,
            .main .block-container,
            div[data-testid="stMain"] .block-container {
                padding-top: 4px !important;
            }
            div[data-testid="stVerticalBlock"] {
                gap: 0.25rem !important;
            }
            div[data-testid="stHorizontalBlock"]:has(.st-key-header_brain_icon) {
                margin-top: -0.5rem !important;
                margin-bottom: -0.2rem !important;
                align-items: center !important;
                min-height: 40px !important;
            }
            .app-header-title {
                transform: translateY(-2px);
                line-height: 1.1;
                margin: 0 !important;
            }
            .app-header-main {
                color: #1e293b;
                font-size: 1.18rem;
                font-weight: 700;
                margin: 0;
            }
            .app-header-subtitle {
                color: #64748b;
                font-size: 0.9rem;
                font-style: italic;
                margin-top: 0.04rem;
            }
            .st-key-header_brain_icon,
            .st-key-main_logout_btn {
                margin-top: -0.3rem !important;
            }
            .st-key-main_logout_btn {
                display: flex !important;
                align-items: center !important;
            }
            .compact-header-divider {
                height: 1px;
                background: #e2e8f0;
                margin: 2px 0 4px;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    header_col, logout_col = st.columns([7, 1.15], vertical_alignment="center")

    with header_col:
        brain_col, title_col = st.columns([0.45, 7.55], vertical_alignment="center")
        with brain_col:
            if st.button("🧠", key="header_brain_icon", help="Click to show/hide helper tips"):
                helper_tab_map = {
                    "💬 Chat": "chat",
                    "📊 Dashboard": "dashboard",
                    "📂 Compare": "compare",
                    "📡 CAPL": "capl"
                }
                current_helper_tab = helper_tab_map.get(st.session_state.get("active_main_tab", "💬 Chat"), "chat")
                state_key = fn._help_state_key(current_helper_tab)
                st.session_state[state_key] = not st.session_state.get(state_key, False)
        with title_col:
            st.markdown(
                """
                <div class="app-header-title">
                    <div class="app-header-main">IntelliDoc AI</div>
                    <div class="app-header-subtitle">Smart Document Assistant</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    with logout_col:
        if st.button("🚶 Logout", key="main_logout_btn"):
            now = datetime.now()
            ist_tz = timezone('Asia/Kolkata')
            ist_time = now.astimezone(ist_tz).strftime("%Y-%m-%d %H:%M:%S %Z")

            # Calculate usage time
            usage_seconds = 0
            if st.session_state.start_time is not None:
                usage_seconds = int(time.time() - st.session_state.start_time)

            hours, remainder = divmod(usage_seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            usage_time_str = f"{hours}h {minutes}m {seconds}s"

            st.session_state.login_history.append({
                "username": st.session_state.logged_in_username,
                "role": st.session_state.user_role,
                "action": "logout",
                "timestamp": ist_time,
                "usage_time": usage_time_str
            })
            active_file = "active_users.json"
            if os.path.exists(active_file):
                with open(active_file, "r") as f:
                    active_users = json.load(f)
                active_users = [u for u in active_users if u["username"] != st.session_state.logged_in_username]
                with open(active_file, "w") as f:
                    json.dump(active_users, f)
            goodbye_user = st.session_state.logged_in_username or "User"

            # Clear all user workspace state so the next login starts fresh.
            st.session_state.is_authenticated = False
            st.session_state.logged_in_username = ""
            st.session_state.user_role = None
            st.session_state.user_session_start_time = None
            st.session_state.start_time = None
            st.session_state.uploaded_files = []
            st.session_state.selected_files = []
            st.session_state.file_texts = {}
            st.session_state.excel_data_by_file = {}
            st.session_state.vector_stores = {}
            st.session_state.ask_messages = []
            st.session_state.extracted_images = {}
            st.session_state.chat_file_selection = []
            st.session_state.chat_summary_downloads = {"images": [], "tables": [], "csv": [], "diagrams": []}
            st.session_state.chat_item_downloads = {"csv": [], "diagrams": []}
            st.session_state.messages = []
            st.session_state.compare_file_selection = []
            st.session_state.compare_result_html = None
            st.session_state.compare_result_excel_bytes = None
            st.session_state.compare_result_files = []
            st.session_state.compare_semantic_summary = None
            st.session_state.file_dropdown = "--Select File--"
            st.session_state.dashboard_chart_type = "Pie Chart"
            st.session_state.dashboard_bar_orientation = "Vertical"
            st.session_state.selected_capl_file = "--Select CAPL file--"
            st.session_state.capl_last_analyzed_file = None
            st.session_state.capl_last_issues = None
            st.session_state.capl_editor_ai_fix = ""
            st.session_state.capl_autonomous_goal = ""
            st.session_state.capl_agent_result = ""
            st.session_state.agent_run_history = []
            st.session_state.last_streamed_assistant_index = None
            st.session_state.mobile_sidebar_visible = False
            st.session_state.llm_task = None
            st.session_state.welcome_shown = False
            st.session_state.behavior_tracker = {
                "chat": {"queries": 0, "actions": []},
                "dashboard": {"queries": 0, "actions": []},
                "compare": {"queries": 0, "actions": []},
                "capl": {"queries": 0, "actions": []}
            }
            for helper_tab in ["chat", "dashboard", "compare", "capl"]:
                st.session_state[fn._help_state_key(helper_tab)] = False
            st.session_state.workspace_memory = {
                "chat": [],
                "agent_runs": [],
                "indexed_files": [],
                "memory_events": [],
                "summary": {},
                "metadata": {},
            }
            st.session_state.workspace_memory_loaded = True
            st.session_state.file_uploader_key = int(st.session_state.get("file_uploader_key", 0)) + 1

            try:
                workspace_db_file = os.path.join(APP_DIR, "workspace_memory.db")
                conn = sqlite3.connect(workspace_db_file, check_same_thread=False)
                cursor = conn.cursor()
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS workspace_meta (
                        meta_key TEXT PRIMARY KEY,
                        meta_value TEXT
                    )
                    """
                )
                cursor.execute(
                    "INSERT OR REPLACE INTO workspace_meta (meta_key, meta_value) VALUES (?, ?)",
                    ("workspace_memory", json.dumps(st.session_state.workspace_memory, default=str))
                )
                conn.commit()
                conn.close()
            except Exception:
                pass

            for cache in [fn.FILE_TEXT_CACHE, fn.VECTOR_STORE_CACHE, fn.EXCEL_DATA_CACHE, fn.EMBEDDINGS_CACHE]:
                cache.clear()
            fn.FILE_HASH_CACHE.clear()
            fn.PREVIEW_TOKENS.clear()
            fn.PREVIEW_STORE.clear()
            save_preview_data()

            st.success(f"Goodbye, {goodbye_user}! You have been logged out.")
            st.rerun()

    st.markdown("<div class='compact-header-divider'></div>", unsafe_allow_html=True)

    render_html_frame(
        """
        <style>
            section.main > div:first-child {
                margin-top: -10px !important;
            }
            @keyframes brainGlow {
                0%, 100% {
                    box-shadow: 0 0 16px rgba(59, 130, 246, 0.45), 0 0 32px rgba(79, 70, 229, 0.22);
                    transform: scale(1);
                }
                25% {
                    box-shadow: 0 0 20px rgba(14, 165, 233, 0.55), 0 0 38px rgba(168, 85, 247, 0.20);
                    transform: scale(1.03);
                }
                50% {
                    box-shadow: 0 0 24px rgba(168, 85, 247, 0.65), 0 0 44px rgba(59, 130, 246, 0.28);
                    transform: scale(1.05);
                }
                75% {
                    box-shadow: 0 0 20px rgba(59, 130, 246, 0.55), 0 0 40px rgba(14, 165, 233, 0.24);
                    transform: scale(1.03);
                }
            }
            .header-brain-icon-large {
                animation: brainGlow 2.2s ease-in-out infinite;
                transition: transform 0.2s ease, box-shadow 0.2s ease, background-color 0.2s ease;
                background: rgba(219, 234, 254, 0.95) !important;
                color: #1d4ed8 !important;
                border: 1px solid rgba(59, 130, 246, 0.35) !important;
                width: 46px !important;
                height: 46px !important;
                min-width: 46px !important;
                min-height: 46px !important;
                padding: 0 !important;
                font-size: 1.55rem !important;
                line-height: 1 !important;
                display: inline-flex !important;
                align-items: center !important;
                justify-content: center !important;
                border-radius: 12px !important;
            }
            .header-brain-icon-large:hover {
                transform: scale(1.08);
                box-shadow: 0 0 30px rgba(79, 70, 229, 0.46), 0 0 52px rgba(14, 165, 233, 0.24);
            }
        </style>
        <script>
            const applyBrainIconStyles = () => {
                const root = window.parent ? window.parent.document : document;
                const buttons = Array.from(root.querySelectorAll('button'));
                buttons.forEach(btn => {
                    const title = (btn.getAttribute('title') || '').trim();
                    const text = (btn.innerText || '').trim().replace(/\\s+/g, '');
                    if (title === 'Click to show/hide helper tips' || text === '🧠') {
                        btn.classList.add('header-brain-icon-large');
                        btn.style.setProperty('font-size', '1.55rem', 'important');
                        btn.style.setProperty('padding', '0', 'important');
                        btn.style.setProperty('min-width', '46px', 'important');
                        btn.style.setProperty('min-height', '46px', 'important');
                        btn.style.setProperty('width', '46px', 'important');
                        btn.style.setProperty('height', '46px', 'important');
                        btn.style.setProperty('display', 'inline-flex', 'important');
                        btn.style.setProperty('align-items', 'center', 'important');
                        btn.style.setProperty('justify-content', 'center', 'important');
                        btn.style.setProperty('overflow', 'visible', 'important');
                        btn.style.setProperty('border-radius', '12px', 'important');
                        btn.style.setProperty('line-height', '1', 'important');
                        btn.style.setProperty('box-sizing', 'border-box', 'important');
                        btn.style.setProperty('transform', 'none', 'important');
                        btn.style.setProperty('background-color', 'rgba(219, 234, 254, 0.95)', 'important');
                        btn.style.setProperty('border', '1px solid rgba(59, 130, 246, 0.35)', 'important');
                        btn.style.setProperty('color', '#1d4ed8', 'important');
                        Array.from(btn.querySelectorAll('*')).forEach(child => {
                            child.style.setProperty('font-size', '1.55rem', 'important');
                            child.style.setProperty('line-height', '1', 'important');
                        });
                    }
                });
            };

            const brainObserver = new MutationObserver(() => applyBrainIconStyles());
            if (window.parent && window.parent.document) {
                brainObserver.observe(window.parent.document.body, { childList: true, subtree: true });
            }
            requestAnimationFrame(applyBrainIconStyles);
            setTimeout(applyBrainIconStyles, 300);
        </script>
        """,
        height=0,
    )

    if not st.session_state.get('welcome_shown', False):
        user = st.session_state.get("logged_in_username", "")
        if user:
            st.toast(f"Welcome back, {user} 🎉", icon="🙋")
        else:
            st.toast("Welcome back 🎉", icon="🙋")
        st.session_state.welcome_shown = True

    render_status_strip()




render_mobile_workspace_controls()

if st.session_state.get("is_authenticated"):
    if st.button("📂 Files / Uploads", key="mobile_show_files_btn"):
        st.session_state.mobile_sidebar_visible = True
        st.rerun()


# -------------------------------
# SESSION STATE INITIALIZATION
# -------------------------------
for key in ["uploaded_files", "selected_files", "file_texts", "excel_data_by_file", "vector_stores", "messages",
            "ask_messages", "extracted_images"]:
    if key not in st.session_state:
        st.session_state[key] = [] if key in ["uploaded_files", "selected_files", "messages", "ask_messages"] else {}

if "capl_last_analyzed_file" not in st.session_state:
    st.session_state.capl_last_analyzed_file = None
if "capl_last_issues" not in st.session_state:
    st.session_state.capl_last_issues = None
if "capl_editor_name" not in st.session_state:
    st.session_state.capl_editor_name = "new_script.can"
if "capl_editor_code" not in st.session_state:
    st.session_state.capl_editor_code = """variables
{

}

on message *
{

}
"""
if "capl_editor_ai_fix" not in st.session_state:
    st.session_state.capl_editor_ai_fix = ""
if "is_authenticated" not in st.session_state:
    st.session_state.is_authenticated = False
if "logged_in_username" not in st.session_state:
    st.session_state.logged_in_username = ""
if "user_role" not in st.session_state:
    st.session_state.user_role = None
if "login_history" not in st.session_state:
    st.session_state.login_history = []
if "file_uploader_key" not in st.session_state:
    st.session_state.file_uploader_key = 0
if "compare_result_html" not in st.session_state:
    st.session_state.compare_result_html = None
if "compare_result_excel_bytes" not in st.session_state:
    st.session_state.compare_result_excel_bytes = None
if "compare_result_files" not in st.session_state:
    st.session_state.compare_result_files = []
if "chat_file_selection" not in st.session_state:
    st.session_state.chat_file_selection = []
if "chat_summary_downloads" not in st.session_state:
    st.session_state.chat_summary_downloads = {"images": [], "tables": [], "csv": [], "diagrams": []}
if "compare_file_selection" not in st.session_state:
    st.session_state.compare_file_selection = []
if "file_dropdown" not in st.session_state:
    st.session_state.file_dropdown = "--Select File--"
if "selected_capl_file" not in st.session_state:
    st.session_state.selected_capl_file = "--Select CAPL file--"
if "llm_task" not in st.session_state:
    st.session_state.llm_task = None
if "user_session_start_time" not in st.session_state:
    st.session_state.user_session_start_time = None
if "start_time" not in st.session_state:
    st.session_state.start_time = None
if "active_main_tab" not in st.session_state:
    st.session_state.active_main_tab = "💬 Chat"
if "workspace_memory" not in st.session_state:
    st.session_state.workspace_memory = {
        "chat": [],
        "agent_runs": [],
        "indexed_files": [],
        "memory_events": [],
        "summary": {},
        "metadata": {},
    }
if "workspace_memory_loaded" not in st.session_state:
    st.session_state.workspace_memory_loaded = False
if "capl_autonomous_goal" not in st.session_state:
    st.session_state.capl_autonomous_goal = ""
if "capl_agent_result" not in st.session_state:
    st.session_state.capl_agent_result = ""
if "agent_run_history" not in st.session_state:
    st.session_state.agent_run_history = []
if "last_streamed_assistant_index" not in st.session_state:
    st.session_state.last_streamed_assistant_index = None
if "compare_semantic_summary" not in st.session_state:
    st.session_state.compare_semantic_summary = None

# -------------------------------
# DOCUMENT PREVIEW FUNCTION
# -------------------------------
# Preview processing helpers:
# These functions support the standalone preview page and also provide extracted
# text/data reused by Chat, Dashboard, Compare, and CAPL when files are selected.
# ===============================================
# LAZY LOADING & PERFORMANCE OPTIMIZATION
# ===============================================




















































































































# -------------------------------
# PREVIEW ROUTE HANDLING
# -------------------------------
# If a preview token is present in the query params, the app short-circuits into
# a dedicated document preview screen instead of rendering the main multi-tab UI.
preview_file_from_url = None
query_params = {}

# Try different methods to get query params (for compatibility across Streamlit versions)
try:
    if hasattr(st, "query_params"):
        query_params = st.query_params
    elif hasattr(st, "experimental_get_query_params"):
        query_params = st.experimental_get_query_params() or {}
    elif hasattr(st, "get_query_params"):
        query_params = st.get_query_params() or {}
    else:
        query_params = {}
except Exception:
    query_params = {}

highlight_term = None
preview_page = None
preview_token = None
if "preview_token" in query_params and query_params["preview_token"]:
    preview_value = query_params["preview_token"]
    if isinstance(preview_value, list):
        preview_value = preview_value[0] if preview_value else ""
    preview_token = str(preview_value)
    token_data = fn.PREVIEW_TOKENS.get(preview_token)
    preview_file_from_url = token_data['file_name'] if token_data else None
    if "highlight" in query_params and query_params["highlight"]:
        highlight_value = query_params["highlight"]
        if isinstance(highlight_value, list):
            highlight_value = highlight_value[0] if highlight_value else ""
        highlight_term = urllib.parse.unquote_plus(str(highlight_value))
    if "page" in query_params and query_params["page"]:
        page_value = query_params["page"]
        if isinstance(page_value, list):
            page_value = page_value[0] if page_value else ""
        try:
            preview_page = int(str(page_value))
        except ValueError:
            preview_page = None

    # If preview_token is present but not found, show error and stop
    if not token_data:
        st.title("Document Preview")
        st.warning("This preview link is no longer available.")
        st.info("Return to the main app and click the eye preview button next to the uploaded file again.")
        st.stop()

if preview_file_from_url:
    preview_entry = fn.PREVIEW_STORE.get(preview_token)
    st.title("📄 Document Preview")
    if preview_entry is not None:
        st.markdown(f"### {preview_entry['name']}")
        st.markdown("---")
        render_professional_document_preview(
            preview_entry['name'],
            file_entry=preview_entry,
            highlight_term=highlight_term,
            highlight_page=preview_page,
        )
    else:
        st.error("Preview file not found in the preview store. Please return to the app and click preview again.")
    st.stop()


# -------------------------------
# -------------------------------
# FILE UPLOAD & MANAGEMENT (SIDEBAR)
# -------------------------------
# Sidebar area:
# This block manages login state, file upload/delete, global file selection, and
# preview launch links. Files selected here become available to the individual tabs.
with st.sidebar:
    if st.session_state.is_authenticated:
        if st.button("Open Workspace", key="mobile_open_workspace_btn", use_container_width=True):
            st.session_state.mobile_sidebar_visible = False
            st.rerun()

        creator_timestamp = None
        if st.session_state.user_role == "creator" and st.session_state.login_history:
            creator_entries = [
                entry for entry in st.session_state.login_history
                if entry.get("username") == st.session_state.logged_in_username and entry.get("role") == "creator" and entry.get("action") == "login"
            ]
            if creator_entries:
                creator_timestamp = creator_entries[-1].get("timestamp")

        st.markdown(
            """
            <style>
                [data-testid="stSidebar"] div[data-testid="column"] {
                    display: flex;
                    align-items: center;
                }
                [data-testid="stSidebar"] div[data-testid="column"] > div {
                    width: 100%;
                }
                .file-box {
                    background: #f8fbff;
                    border: 1px solid #d7e3f4;
                    border-radius: 12px;
                    margin-bottom: 8px;
                    color: #173152;
                    font-size: 14px;
                    overflow-wrap: anywhere;
                }
                [data-testid="stSidebar"] [class*="st-key-select_file_"] button[kind="secondary"],
                [data-testid="stSidebar"] [class*="st-key-select_file_"] button[kind="primary"] {
                    width: 100%;
                    min-height: 72px;
                    border-radius: 12px;
                    padding: 12px 14px;
                    font-size: 14px;
                    text-align: left;
                    justify-content: flex-start;
                    white-space: normal;
                    line-height: 1.4;
                }
                [data-testid="stSidebar"] [class*="st-key-select_file_"] button[kind="secondary"] {
                    background: #f5fbff;
                    border: 1px solid #d0e8f8;
                    color: #1e293b;
                }
                [data-testid="stSidebar"] [class*="st-key-select_file_"] button[kind="primary"] {
                    background: #ffe7d6 !important;
                    border: 2px solid #ffbea3 !important;
                    color: #5f351c !important;
                    box-shadow: inset 0 0 0 1px rgba(255, 190, 163, 0.55);
                }
                [data-testid="stSidebar"] [class*="st-key-select_file_"] button[kind="primary"]:hover {
                    background: #ffd3b5 !important;
                    border-color: #ffb18f !important;
                }
                [data-testid="stSidebar"] div[data-testid="stLinkButton"] a {
                    min-height: 38px;
                    border-radius: 12px;
                    padding: 0.35rem 0.5rem;
                    white-space: nowrap;
                    justify-content: center;
                    background: #f5fbff !important;
                    border: 1px solid #d0e8f8 !important;
                    color: #173152 !important;
                }
                [data-testid="stSidebar"] div[data-testid="stLinkButton"] a:hover {
                    background: #e8f6ff !important;
                    border-color: #a0c8e8 !important;
                }
                [data-testid="stSidebar"] [class*="st-key-del_file_"] button[kind="tertiary"] {
                    background: transparent;
                    border: none;
                    box-shadow: none;
                    color: #6c7280;
                    min-height: 38px;
                }
                [data-testid="stSidebar"] [class*="st-key-del_file_"] button[kind="tertiary"]:hover {
                    background: #f3f6fb;
                    color: #b42318;
                }
            </style>
            """,
            unsafe_allow_html=True
        )

        # Mercedes Logo in Sidebar
        if logo_data:
            st.markdown(
                f'''
                <div style="text-align: center; margin-bottom: 20px;">
                    <img src="data:image/gif;base64,{logo_data}" style="width: 40px; height: 40px;">
                    <div style="font-size: 15px; color: var(--text-secondary); margin-top: 4px; font-weight: 500;">Mercedes-Benz</div>
                </div>
                ''',
                unsafe_allow_html=True,
            )
        
        st.header("Upload Documents")
        st.info("1) Upload files." \
        " 2) Click the file cards you need. " \
        "3) Switch tabs and work with selected files.")
        new_files = st.file_uploader(
            "Upload PDF, Word, PPT, Excel, CSV, TXT, HTML, ODT, RTF, Pages, CAPL, Images",
            type=["pdf", "doc", "docx", "txt", "md", "log", "ppt", "pptx", "xls", "xlsx", "csv", "html", "htm", "odt", "rtf", "pages", "capl", "can", "png", "jpg", "jpeg", "gif", "bmp", "webp"],
            accept_multiple_files=True,
            key=f"file_uploader_{st.session_state.file_uploader_key}"
        )

        if new_files:
            existing_names = {f["name"] for f in st.session_state.uploaded_files}
            new_file_names = []
            for file in new_files:
                if file.name not in existing_names:
                    file_bytes = file.read()
                    st.session_state.uploaded_files.append({
                        "name": file.name,
                        "bytes": file_bytes,
                        "status": "queued",
                    })
                    new_file_names.append(file.name)
                    existing_names.add(file.name)

            if new_file_names:
                for file_name in new_file_names:
                    update_uploaded_file_status(file_name, "processing")
                st.toast("📄 File uploaded — initializing pipeline...")
                st.info("📄 File uploaded — processing in background. You can select files and keep working.")
                st.session_state.workspace_memory["indexed_files"] = sorted(set(st.session_state.workspace_memory.get("indexed_files", []) + new_file_names))
                record_workspace_memory_event(
                    "upload",
                    "Documents queued for shared memory",
                    "Uploaded and queued for processing: " + ", ".join(new_file_names),
                    source="Upload",
                )
                save_workspace_memory()
                save_memory_log("upload", f"Queued {len(new_file_names)} new file(s)", {"files": new_file_names})
                st.session_state.messages = []
                st.session_state.chat_summary_downloads = {"images": [], "tables": [], "csv": [], "diagrams": []}
                st.session_state.chat_file_selection = []
                st.success("✅ New files uploaded. Chat history has been cleared.")

        st.markdown("---")
        st.markdown("### Uploaded files")
        
        for idx, file_dict in enumerate(st.session_state.uploaded_files[:]):
            cols = st.columns([0.56, 0.27, 0.17], vertical_alignment="center")
            with cols[0]:
                file_name = file_dict["name"]
                file_status = file_dict.get("status", "ready")
                if file_status == "ready":
                    status_label = "ready"
                elif file_status == "queued":
                    status_label = "queued..."
                else:
                    status_label = f"{file_status}..."
                is_selected = file_name in st.session_state.selected_files
                file_label = f"{file_name} - {status_label}"
                button_label = file_label if not is_selected else f"Selected: {file_label}"
                if st.button(
                    button_label,
                    key=f"select_file_{idx}",
                    help=f"Click to {'remove' if is_selected else 'add'} {file_name}",
                    use_container_width=True,
                    type="primary" if is_selected else "secondary",
                ):
                    if is_selected:
                        st.session_state.selected_files.remove(file_name)
                    else:
                        st.session_state.selected_files.append(file_name)
                    st.rerun()
            with cols[1]:
                preview_url = create_preview_link(file_name)
                st.link_button(
                    "Open",
                    preview_url or "#",
                    help=f"Open preview for {file_name}",
                    icon="👁️",
                    type="secondary",
                    use_container_width=True,
                    disabled=preview_url is None,
                )
            with cols[2]:
                if st.button("🗑️", key=f"del_file_{idx}", help=f"Delete {file_name}", use_container_width=True, type="tertiary"):
                    deleted_name = file_name
                    st.session_state.uploaded_files = [f for f in st.session_state.uploaded_files if
                                                       f["name"] != deleted_name]

                    if deleted_name in st.session_state.selected_files:
                        st.session_state.selected_files.remove(deleted_name)

                    for key in ["file_texts", "excel_data_by_file", "vector_stores"]:
                        st.session_state.get(key, {}).pop(deleted_name, None)

                    if "workspace_memory" in st.session_state:
                        st.session_state.workspace_memory["indexed_files"] = [
                            f for f in st.session_state.workspace_memory.get("indexed_files", [])
                            if f != deleted_name
                        ]

                    if st.session_state.capl_last_analyzed_file == deleted_name:
                        st.session_state.capl_last_analyzed_file = None
                        st.session_state.capl_last_issues = None

                    st.session_state.file_uploader_key = int(st.session_state.get("file_uploader_key", 0)) + 1
                    st.rerun()
        st.markdown("*Selected files above are available across all tabs.*")
        st.markdown("---")
        if st.button("Clear All Files"):
            for key in ["uploaded_files", "selected_files", "file_texts", "excel_data_by_file", "vector_stores",
                        "messages"]:
                st.session_state[key].clear()
            st.session_state.chat_summary_downloads = {"images": [], "tables": [], "csv": [], "diagrams": []}
            st.session_state.chat_file_selection = []
            st.session_state.capl_last_analyzed_file = None
            st.session_state.capl_last_issues = None
            st.session_state.file_uploader_key += 1
            st.rerun()

        if st.session_state.user_role == "creator":
            st.markdown("---")
            st.subheader("Creator Login / Logout History")
            if st.session_state.login_history:
                st.table(pd.DataFrame(st.session_state.login_history))
            st.markdown("---")
            st.subheader("User Statistics")
            total_opened = len(set(h["username"] for h in st.session_state.login_history))
            active_file = "active_users.json"
            if os.path.exists(active_file):
                with open(active_file, "r") as f:
                    active_users = json.load(f)
                now = datetime.now()
                current_active = len(set(u["username"] for u in active_users if
                                         datetime.fromisoformat(u["timestamp"]) > now - timedelta(minutes=30)))
            else:
                current_active = 0
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total Users Opened", total_opened)
            with col2:
                st.metric("Currently Active Users", current_active)
            st.markdown("---")
            if st.button("Clean Old Active Users"):
                active_file = "active_users.json"
                if os.path.exists(active_file):
                    with open(active_file, "r") as f:
                        active_users = json.load(f)
                    now = datetime.now()
                    active_users = [u for u in active_users if
                                    datetime.fromisoformat(u["timestamp"]) > now - timedelta(hours=1)]
                    with open(active_file, "w") as f:
                        json.dump(active_users, f)
                    st.success("Cleaned old active users.")
                else:
                    st.info("No active users file found.")


# -------------------------------
# TEXT EXTRACTION
# -------------------------------
# Excel extraction helper:
# Primarily used by Dashboard previews/charts, but cached so other tabs can reuse
# the parsed spreadsheet rows without re-reading the uploaded file each rerun.


# -------------------------------
# PROCESS FILES & BUILD VECTOR STORES
# -------------------------------
# AI/vector helpers:
# These are mainly used by the Chat tab for semantic retrieval and LLM answers.
# They also centralize file preprocessing so each tab can rely on the same cache.








WORKSPACE_DB_FILE = os.path.join(APP_DIR, "workspace_memory.db")
WORKSPACE_MEMORY_KEY = "workspace_memory"
































































ensure_workspace_memory_loaded()




SUMMARY_STOPWORDS = {
    "the", "and", "for", "with", "that", "this", "from", "are", "was", "were", "into", "your", "have",
    "has", "had", "not", "but", "you", "all", "can", "will", "use", "using", "used", "how", "what",
    "when", "where", "which", "while", "into", "more", "than", "their", "there", "about", "after",
    "before", "within", "without", "each", "page", "pages", "table", "tables", "image", "images",
    "document", "content", "metadata", "information", "product", "file", "text"
}






# Chat summary helper:
# Used only in the Chat tab after summarize/analyze actions to expose extracted
# images and tables as downloads for the files currently selected in chat.




































# Document structure helpers:
# Used by the Chat "overview" flow and preview links to identify headings,
# table-of-contents entries, and likely page numbers inside uploaded documents.




















CREATOR_USERNAME = "Vignesh"
CREATOR_PASSWORD = "Rider@100"

# Login gate:
# This runs before the main app tabs are shown. It keeps the creator/user access
# flow in one place so authentication checks do not have to be repeated per tab.
# ================================
# PREMIUM LOGIN EXPERIENCE
# Replace your current login gate block with this
# ================================
if not st.session_state.is_authenticated and "preview_token" not in query_params:
    # Custom CSS for the login page
    st.markdown("""
    <style>
        html, body { margin: 0; padding: 0; height: 100%; }
        .block-container {
            max-width: none !important;
            padding: clamp(20px, 3vw, 36px) !important;
            margin: 0 !important;
            width: 100vw !important;
            min-height: 100vh !important;
        }
        .main {
            width: 100% !important;
            padding: 0 !important;
            min-height: 100vh !important;
        }

        /* Full viewport container */
        [data-testid="stAppViewContainer"] {
            background:
                radial-gradient(circle at top left, rgba(176, 224, 230, 0.35) 0%, rgba(176, 224, 230, 0) 40%),
                radial-gradient(circle at top right, rgba(230, 245, 250, 0.3) 0%, rgba(230, 245, 250, 0) 45%),
                radial-gradient(circle at bottom left, rgba(176, 196, 222, 0.25) 0%, rgba(176, 196, 222, 0) 40%),
                linear-gradient(135deg, #ffffff 0%, #f8fafc 40%, #eef2f7 100%) !important;  # login background
            min-height: 100vh !important;
            display: flex !important;
            align-items: stretch !important;
        }

        /* Override Streamlit's default column behavior */
        [data-testid="column"]:first-child {
            flex: 1 !important;
            padding: clamp(36px, 4vw, 64px) clamp(32px, 5vw, 84px) !important;
            display: flex !important;
            flex-direction: column !important;
            justify-content: center !important;
            color: #3B5E7F !important;
            min-height: calc(100vh - 2 * clamp(20px, 3vw, 36px)) !important;
        }
        [data-testid="column"]:nth-child(2) {
            width: min(460px, 100%) !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            padding: clamp(20px, 2.6vw, 32px) !important;
            flex-shrink: 0 !important;
            min-height: calc(100vh - 2 * clamp(20px, 3vw, 36px)) !important;
        }

        /* Branding elements */
        .brand-strip {
            display: flex !important;
            align-items: center !important;
            gap: 12px !important;
            #margin-bottom: 32px !important;
           perspective: 800px;      
        }
        .login-panel {
            width: min(440px, 100%);
            padding: 36px 32px 30px;
            border-radius: 24px;
            background: linear-gradient(180deg, rgba(255, 253, 250, 0.95) 0%, rgba(245, 250, 252, 0.92) 100%);
            border: 1.5px solid rgba(176, 224, 230, 0.45);
            box-shadow: 0 20px 50px rgba(135, 206, 235, 0.1), 0 0 30px rgba(176, 196, 222, 0.08);
            backdrop-filter: blur(12px);
        }
        .brand-logo-3d {
            width: 60px !important;
            height: 60 px !important;
            perspective: 800px;    
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            animation: float 3s ease-in-out infinite;       
        }  
        @keyframes float {
           0%, 100% { transform: translateY(0px); }
          50% { transform: translateY(-4px); }
        }  
        .logo-inner{
                width : 100%;
                height: 100%;
                border-radius:50%;
                background: radial-gradient(circle at 30% 30%, #dcdcdc, #a0a0a0);
                display: flex;
                align-items: center;
                justify-content: center;
                transform-style: preserve-3d;
                transition: transform 0.2s ease;
                box-shadow: 0 8px 20px rgba(0,0,0,0.15);
            }   
        .star{
              font-size: 28px;
              color: white;
              transform: translateZ(10px); 
            }                      
        .brand-label {
            font-size: 0.82rem !important;
            color: rgba(91, 127, 166, 0.9) !important;
            letter-spacing: 0.18em !important;
            text-transform: uppercase !important;
        }
        .ai-branding {
            display : flex;
            gap : 8px;
            align-items : baseline;
            margin-bottom: 28px!important;
        }
        
        .ai-title {
           font-size: clamp(2rem, 3.5vw, 2.8rem);
           font-weight: 800;
           color: #2C5F7F;
        }

        .ai-subtitle {
         font-size: 14px;
         font-style: italic;
         font-weight: 400;
         color: rgba(44, 95, 127, 0.6);
        }
        .ai-tagline {
            font-size: clamp(2.2rem, 4vw, 3.2rem) !important;
            font-weight: 800 !important;
            line-height: 1.15 !important;
            color: #3B5E7F !important;
            margin-bottom: 32px !important;
            letter-spacing: -0.03em !important;
        }
        .ai-description {
            font-size: 1.02rem !important;
            line-height: 1.7 !important;
            color: rgba(91, 127, 166, 0.9) !important;
            margin-bottom: 48px !important;
        }
        .trust-row {
            color: rgba(91, 127, 166, 0.8) !important;
            font-size: 0.95rem !important;
            margin-bottom: 56px !important;
            line-height: 1.6 !important;
        }
        /* Feature cards */
        .feature-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(280px, 1fr));
            gap: 16px;
            align-items: stretch;
            margin-top: 8px;
        }
        .feature-card {
            background: rgba(176, 224, 230, 0.15);
            border: 1px solid rgba(176, 196, 222, 0.3);
            border-radius: 12px;
            color: #3B5E7F;
            padding: 18px 18px 16px;
            min-height: 150px;
            display: flex;
            flex-direction: column;
            justify-content: flex-start;
            backdrop-filter: blur(4px);
        }
        .feature-card h4 {
            margin: 0 0 14px 0;
            color: #0d5aa7;
            font-size: 1.08rem;
            font-weight: 700;
        }
        .feature-card ul {
            margin: 0;
            padding-left: 20px;
        }
        .feature-card li {
            margin: 0 0 8px 0;
            line-height: 1.45;
            color: #0d5aa7;
        }
        .login-panel .stButton button {
            background: linear-gradient(135deg, #87CEEB 0%, #6BA3C5 100%) !important;
            color: #ffffff !important;
            border: 0 !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
            width: 100% !important;
            padding: 12px 16px !important;
            margin-top: 8px !important;
        }

        .login-panel .stButton button:hover {
            background: linear-gradient(135deg, #ADD8E6 0%, #87CEEB 100%) !important;
            transform: translateY(-1px) !important;
        }

        .login-note {
            color: rgba(100, 100, 140, 0.7) !important;
            font-size: 0.85rem !important;
            text-align: center !important;
            margin-top: 16px !important;
        }

        /* Responsive design */
        @media (max-width: 768px) {
            [data-testid="column"]:first-child {
                padding: 24px 12px 16px !important;
                order: 2 !important;
                min-height: auto !important;
            }
            [data-testid="column"]:nth-child(2) {
                width: 100% !important;
                padding: 12px !important;
                order: 1 !important;
                min-height: auto !important;
            }
            .ai-tagline {
                font-size: 2rem !important;
            }
            .login-panel {
                width: 100% !important;
                padding: 24px 20px !important;
            }
            .feature-grid {
                grid-template-columns: 1fr;
            }
        }
    </style>
    """, unsafe_allow_html=True)

    # Create the new flexbox layout using Streamlit columns - Login page layout
    left_col, right_col = st.columns([3, 1.3])

    # Apply additional form styling for login elements
    st.markdown("""
    <style>
        /* Login heading and subheading */
        .login-panel .login-heading {
            margin-top: 0 !important;
        }
        .login-heading {
            font-size: clamp(1.8rem, 4vw, 2.8rem) !important;
            color: #2C5F7F !important;
            font-weight: 700 !important;
            margin-bottom: 10px !important;
            line-height: 1.1 !important;
        }

        .login-subheading {
            font-size: clamp(0.95rem, 2vw, 1.05rem) !important;
            color: #5B7FA6 !important;
            font-weight: 500 !important;
            margin-bottom: 20px !important;
        }

        /* Form elements */
        .login-panel [data-testid="stTextInput"] {
            margin-bottom: 8px !important;
        }
        .login-panel [data-testid="stTextInput"] label,
        .login-panel [data-testid="stTextInput"] label p,
        .login-panel [data-testid="stTextInput"] label span,
        .login-panel [data-testid="stTextInput"] p {
            color: #3B5E7F !important;
            -webkit-text-fill-color: #3B5E7F !important;
            opacity: 1 !important;
        }
        .login-panel [data-baseweb="base-input"],
        .login-panel [data-baseweb="input"],
        .login-panel [data-testid="stTextInput"] > div > div {
            background: rgba(230, 244, 248, 0.85) !important;
            border: 1.5px solid rgba(176, 224, 230, 0.6) !important;
            border-radius: 12px !important;
            box-shadow: none !important;
        }
        .login-panel [data-baseweb="base-input"]:focus-within,
        .login-panel [data-baseweb="input"]:focus-within,
        .login-panel [data-testid="stTextInput"] > div > div:focus-within {
            border: 1.5px solid #87CEEB !important;
            box-shadow: 0 0 0 2px rgba(135, 206, 235, 0.3) !important;
                outline : none !important;
        }
                    
         /*Remove streamlit red focus */
        .login-panel input:focus,
        .login-panel input:invalid {
                outline:none !important;
                box-shadow:none ! important;
                border-color: #87CEEB !important;
            }   
                                        
        .login-panel input[type="text"],
        .login-panel input[type="password"] {
            background: transparent !important;
            color: #2C5F7F !important;
            caret-color: #4D94B9 !important;
            -webkit-text-fill-color: #2C5F7F !important;
            border: none !important;
            box-shadow: none !important;
            font-weight: 600 !important;
            letter-spacing: 0.2px !important;
            font-size: 1rem !important;
            padding: 12px 16px !important;
        }
        .login-panel input[type="text"]::placeholder,
        .login-panel input[type="password"]::placeholder {
            color: rgba(91, 127, 166, 0.65) !important;
            opacity: 1 !important;
            -webkit-text-fill-color: rgba(91, 127, 166, 0.65) !important;
        }
        .login-panel input[type="text"]:focus,
        .login-panel input[type="password"]:focus {
            background: transparent !important;
            color: #2C5F7F !important;
            -webkit-text-fill-color: #2C5F7F !important;
            outline: none !important;
        }

        .login-note {
            color: rgba(91, 127, 166, 0.75) !important;
            font-size: 0.85rem !important;
            text-align: center !important;
            margin-top: 16px !important;
        }
        .login-panel [data-testid="stCaptionContainer"],
        .login-panel [data-testid="stCaptionContainer"] p,
        .login-panel .stCaption,
        .login-panel .stCaption p {
            color: rgba(91, 127, 166, 0.85) !important;
            -webkit-text-fill-color: rgba(91, 127, 166, 0.85) !important;
            opacity: 1 !important;
            font-size: 0.84rem !important;
        }
        .login-panel .stButton > button,
        .login-panel div.stButton > button {
            width: 100% !important;
            min-height: 42px !important;
            padding: 0.45rem 0.9rem !important;
            font-size: 0.98rem !important;
            border-radius: 10px !important;
        }

        /* Responsive design */
        @media (max-width: 768px) {
            [data-testid="column"]:first-child {
                padding: 24px 12px 16px !important;
                order: 2 !important;
            }
            [data-testid="column"]:nth-child(2) {
                width: 100% !important;
                padding: 12px !important;
                order: 1 !important;
            }
            .ai-tagline {
                font-size: 2rem !important;
            }
            .login-panel {
                width: 100% !important;
                padding: 24px 20px !important;
            }
        }
         /* 🔥 HARD OVERRIDE - kills ALL red focus from Streamlit/BaseWeb */

            .login-panel *:focus {
             outline: none !important;
            }  
            .login-panel input,
            .login-panel textarea {
               box-shadow: none !important;
            }

           /* Target BaseWeb internal input container */
            .login-panel div[data-baseweb="base-input"] {
              border: 1.5px solid rgba(176, 224, 230, 0.6) !important;
            }   

            /* Focus state */
           .login-panel div[data-baseweb="base-input"]:focus-within {
               border: 1.5px solid #87CEEB !important;
              box-shadow: 0 0 0 2px rgba(135, 206, 235, 0.3) !important;
            }    

            /* Remove error (red) state completely */
            .login-panel div[data-baseweb="base-input"][aria-invalid="true"] {
                border: 1.5px solid #87CEEB !important;
                box-shadow: none !important;
            }

            /* Also kill any red glow from deeper layers */
            .login-panel div[data-baseweb="input"] {
               box-shadow: none !important;
            }   

         /* 🚨 NUCLEAR OVERRIDE — removes ALL red states */

            .login-panel *[aria-invalid="true"],
            .login-panel *[data-baseweb="base-input"][aria-invalid="true"],
            .login-panel *[data-baseweb="input"][aria-invalid="true"] {
                border-color: #87CEEB !important;
                box-shadow: none !important;
            }

         /* Remove any red focus ring from BaseWeb */
            .login-panel *:focus-visible {
                outline: none !important;
                box-shadow: none !important;
            }

            /* Force our focus style ONLY */
            .login-panel div[data-baseweb="base-input"]:focus-within {
                border: 1.5px solid #87CEEB !important;
              box-shadow: 0 0 0 2px rgba(135, 206, 235, 0.3) !important;
            }

            /* Kill ANY red borders globally inside login */
            .login-panel input,
            .login-panel div,
            .login-panel textarea {
             border-color: rgba(176, 224, 230, 0.6) !important;
            } 
            /* 🚨 HARD KILL: BaseWeb invalid (red) state */

            .login-panel div[data-baseweb="base-input"][aria-invalid="true"],
            .login-panel div[data-baseweb="input"][aria-invalid="true"],
            .login-panel [data-baseweb="base-input"] {
               border: 1.5px solid rgba(176, 224, 230, 0.6) !important;
               box-shadow: none !important;
            }

            /* Override inner input focus ring */
            .login-panel input {
               outline: none !important;
               box-shadow: none !important;
            }

            /* Force consistent focus (blue theme only) */
            .login-panel div[data-baseweb="base-input"]:focus-within {
              border: 1.5px solid #87CEEB !important;
             box-shadow: 0 0 0 2px rgba(135, 206, 235, 0.25) !important;
            }         

    </style>
    """, unsafe_allow_html=True)

    with left_col:
        if logo_data:
            logo_display = f'''
            <img src = "data:image/gif;base64,{logo_data}"
            style = "width: 36px; height: 36px; object-fit;contain;">'''
        else:
             logo_display = '<div class = "star">★</div>'
        st.markdown(f"""
                    <div class = "brand-strip">
                        <div class = " brand-logo-3d">
                           {logo_display}
                        </div>
                    </div>
                    <div class = "brand-label">Mercedes_Benz</div>
                </div>
                """,unsafe_allow_html=True)
            
        
        
        st.markdown("""<div class="ai-branding"> 
        <span class= "ai-title"> IntelliDoc AI </span> 
        <span class="ai-subtitle">-Smart Document Assistant </span>
        </div>""", unsafe_allow_html=True)
        st.markdown('<h1 class="ai-tagline">Where Documents Become Intelligence</h1>', unsafe_allow_html=True)
        st.markdown('<p class="ai-description">for secure document insight, comparison, dashboards, and automation.</p>', unsafe_allow_html=True)
        st.markdown(
            """
            <div class="feature-grid">
                <div class="feature-card">
                    <h4>💬 Chat</h4>
                    <ul>
                        <li>Ask questions about uploaded files</li>
                        <li>Context-aware responses</li>
                        <li>Multi-file semantic understanding</li>
                    </ul>
                </div>
                <div class="feature-card">
                    <h4>📊 Dashboard</h4>
                    <ul>
                        <li>Excel/CSV visualization</li>
                        <li>Export insights</li>
                    </ul>
                </div>
                <div class="feature-card">
                    <h4>🔄 Compare</h4>
                    <ul>
                        <li>Compare 2+ files</li>
                        <li>Word-level diff</li>
                        <li>Inline visual comparison</li>
                        <li>Export results to Excel</li>
                    </ul>
                </div>
                <div class="feature-card">
                    <h4>📡 CAPL</h4>
                    <ul>
                        <li>Upload or create <code>.can</code> files</li>
                        <li>Built-in CAPL editor</li>
                        <li>Code analysis &amp; issue detection</li>
                        <li>Suggestions &amp; improvements</li>
                    </ul>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with right_col:  # type: ignore  # right_col is defined at line 3699
        #st.markdown('<div class="login-panel">', unsafe_allow_html=True)
        #st.markdown('<div class="login-heading">Welcome back</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-subheading">Sign in to IntelliDoc AI</div>', unsafe_allow_html=True)

        login_username = st.text_input("👤 Username", placeholder="Username", key="username")
        login_password = st.text_input("🔒 Password", type="password", placeholder="Password", key="password")

        st.caption("Standard users can leave password empty")

        access_clicked = st.button("Access", use_container_width=False, key="signin")

        st.markdown('</div>', unsafe_allow_html=True)

    if access_clicked:
        cleaned_username = (login_username or "").strip()
        cleaned_password = (login_password or "").strip()

        if cleaned_username == CREATOR_USERNAME and cleaned_password == CREATOR_PASSWORD:
            st.session_state.is_authenticated = True
            st.session_state.logged_in_username = cleaned_username
            st.session_state.user_role = "creator"
            st.session_state.user_session_start_time = datetime.now().isoformat()
            st.session_state.start_time = time.time()

            ist_tz = timezone('Asia/Kolkata')
            ist_time = datetime.now(ist_tz).strftime("%Y-%m-%d %H:%M:%S %Z")
            st.session_state.login_history.append({
                "username": cleaned_username,
                "role": "creator",
                "action": "login",
                "timestamp": ist_time,
                "usage_time": "-"
            })

            active_file = "active_users.json"
            now = datetime.now()
            if os.path.exists(active_file):
                with open(active_file, "r") as f:
                    active_users = json.load(f)
            else:
                active_users = []

            active_users = [u for u in active_users if u.get("username") != cleaned_username]
            active_users.append({"username": cleaned_username, "timestamp": now.isoformat()})

            with open(active_file, "w") as f:
                json.dump(active_users, f)

            st.success("✅ Creator access granted")
            st.rerun()

        elif cleaned_username and len(cleaned_username) > 3 and cleaned_password == "":
            st.session_state.is_authenticated = True
            st.session_state.logged_in_username = cleaned_username
            st.session_state.user_role = "user"
            st.session_state.user_session_start_time = datetime.now().isoformat()
            st.session_state.start_time = time.time()

            ist_tz = timezone('Asia/Kolkata')
            ist_time = datetime.now(ist_tz).strftime("%Y-%m-%d %H:%M:%S %Z")
            st.session_state.login_history.append({
                "username": cleaned_username,
                "role": "user",
                "action": "login",
                "timestamp": ist_time,
                "usage_time": "-"
            })

            active_file = "active_users.json"
            now = datetime.now()
            if os.path.exists(active_file):
                with open(active_file, "r") as f:
                    active_users = json.load(f)
            else:
                active_users = []

            active_users = [u for u in active_users if u.get("username") != cleaned_username]
            active_users.append({"username": cleaned_username, "timestamp": now.isoformat()})

            with open(active_file, "w") as f:
                json.dump(active_users, f)

            st.success(f"✅ Welcome, {cleaned_username}!")
            st.rerun()

        else:
            st.error("❌ Invalid credentials. Creator needs password. Users need username >3 chars with empty password.")

    st.stop()


# Files will be processed on-demand per tab when selected
# -------------------------------
# HELPER CHART FUNCTIONS
# -------------------------------
# Dashboard chart helpers:
# These are used by the Dashboard tab when an XLSX/HTML file is selected and the
# user wants counts shown as bar or pie charts.






# -------------------------------
# INLINE MULTI-FILE DIFF (HTML)
# -------------------------------
# Compare tab HTML diff helper:
# Generates the inline visual comparison shown in the Compare tab and also reused
# from Chat when the user asks to compare multiple selected documents.








# -------------------------------
# COMPARE EXCEL HIGHLIGHT
# -------------------------------
# Compare tab Excel export helper:
# Builds the downloadable workbook used in the Compare tab so users can inspect
# mismatches outside the Streamlit UI.




# -------------------------------
# CAPL Complier
# -------------------------------
# CAPL analyzer helpers:
# These functions are used only by the CAPL tab for syntax checking, issue
# listing, and highlighted code rendering inside the CAPL editor/viewer panels.




# -------------------------------
# CAPL CODE DETECTION
# -------------------------------










# Shared UI helpers:
# These small functions are reused across multiple tabs to show the current
# sidebar selection, tab-level file context, and floating helper popups.






# Floating icon function removed - helper is now triggered by header 🧠 icon




# Define keywords for each tab
tab_keywords = {
    "chat": ["memory", "overview", "summary", "count", "find", "analyze", "details", "downloads"],
    "dashboard": ["insights", "memory", "themes", "entities", "risks", "charts", "metrics", "reports"],
    "compare": ["semantic", "differences", "compare", "changes", "diff", "side-by-side", "inline", "excel"],
    "capl": ["agents", "syntax", "variables", "errors", "debug", "code", "fix", "run history"]
}












# -------------------------------
# TABS
# -------------------------------
# Main application tabs:
# Each block below owns one visible area of the app. If you want to change a
# feature, start in the matching tab block and then follow the helper comments above.

# Session-backed main navigation:
# Premium glass navigation with one shared indicator/glow state.
st.markdown("""
    <style>
    .st-key-active_main_tab {
        margin-top: 16px !important;
        margin-bottom: 8px !important;
        position: relative;
        --ai-nav-accent: var(--accent, #3b82f6);
        --ai-nav-accent-rgb: 59, 130, 246;
        --ai-nav-muted: #64748b;
        --ai-nav-text: #0f172a;
        --ai-nav-indicator-x: 0px;
        --ai-nav-indicator-width: 56px;
    }

    .st-key-active_main_tab div[role="radiogroup"] > label > div:first-child {
        display: none !important;
    }

    .st-key-active_main_tab div[role="radiogroup"] {
        align-items: center;
        border-bottom: 1px solid rgba(var(--ai-nav-accent-rgb), 0.10);
        display: flex;
        gap: clamp(18px, 4vw, 36px);
        justify-content: center;
        margin: 0 !important;
        overflow: visible;
        padding: 0 4px 12px !important;
        position: relative;
    }

    .st-key-active_main_tab div[role="radiogroup"] > label {
        align-items: center !important;
        background: transparent !important;
        border: none !important;
        border-radius: 0 !important;
        box-shadow: none !important;
        color: var(--ai-nav-muted) !important;
        cursor: pointer;
        display: inline-flex !important;
        font-weight: 650;
        gap: 7px;
        height: auto;
        isolation: isolate;
        line-height: 1.2 !important;
        min-height: auto !important;
        overflow: visible;
        padding: 9px 18px !important;
        position: relative;
        transform: translateZ(0);
        transition:
            color 320ms cubic-bezier(0.22, 1, 0.36, 1),
            transform 320ms cubic-bezier(0.22, 1, 0.36, 1),
            opacity 320ms cubic-bezier(0.22, 1, 0.36, 1);
        white-space: nowrap;
        z-index: 2;
    }

    .st-key-active_main_tab div[role="radiogroup"] > label::before {
        background:
            radial-gradient(circle,
                rgba(var(--ai-nav-accent-rgb), 0.24) 0%,
                rgba(var(--ai-nav-accent-rgb), 0.16) 34%,
                rgba(var(--ai-nav-accent-rgb), 0.06) 58%,
                rgba(var(--ai-nav-accent-rgb), 0) 74%);
        border-radius: 999px;
        content: "";
        filter: blur(7px);
        height: 34px;
        left: 9px;
        opacity: 0;
        pointer-events: none;
        position: absolute;
        top: 50%;
        transform: translateY(-50%) scale(0.92);
        transition: opacity 360ms cubic-bezier(0.22, 1, 0.36, 1);
        width: 34px;
        z-index: -1;
    }

    .st-key-active_main_tab div[role="radiogroup"] > label:hover {
        color: #334155 !important;
        transform: scale(1.035);
    }

    .st-key-active_main_tab div[role="radiogroup"] > label[data-checked="true"],
    .st-key-active_main_tab div[role="radiogroup"] > label.ai-nav-active {
        background: transparent !important;
        border-radius: 0 !important;
        border: none !important;
        box-shadow: none !important;
        color: var(--ai-nav-text) !important;
        font-weight: 800;
    }

    .st-key-active_main_tab div[role="radiogroup"] > label[data-checked="true"]::before,
    .st-key-active_main_tab div[role="radiogroup"] > label.ai-nav-active::before {
        animation: aiNavGlowBreath 4.6s cubic-bezier(0.45, 0, 0.2, 1) infinite;
        opacity: 0.55;
    }

    .st-key-active_main_tab div[role="radiogroup"] > label p {
        color: inherit !important;
        font: inherit !important;
        line-height: inherit !important;
        margin: 0 !important;
        position: relative;
        z-index: 1;
    }

    .st-key-active_main_tab div[role="radiogroup"] > label::after {
        display: none !important;
    }

    .st-key-active_main_tab .ai-nav-indicator {
        background:
            linear-gradient(90deg,
                rgba(var(--ai-nav-accent-rgb), 0.08),
                rgba(var(--ai-nav-accent-rgb), 0.62) 36%,
                rgba(255, 255, 255, 0.88) 50%,
                rgba(var(--ai-nav-accent-rgb), 0.62) 64%,
                rgba(var(--ai-nav-accent-rgb), 0.08));
        border-radius: 999px;
        bottom: -1px;
        box-shadow:
            0 0 9px rgba(var(--ai-nav-accent-rgb), 0.34),
            0 0 18px rgba(var(--ai-nav-accent-rgb), 0.18);
        content: "";
        display: block;
        filter: blur(0.25px);
        height: 3px;
        left: 0;
        opacity: 0.96;
        pointer-events: none;
        position: absolute;
        transform: translate3d(var(--ai-nav-indicator-x), 0, 0);
        transition:
            transform 480ms cubic-bezier(0.2, 1.18, 0.32, 1),
            width 480ms cubic-bezier(0.2, 1.18, 0.32, 1),
            opacity 240ms ease;
        width: var(--ai-nav-indicator-width);
        z-index: 1;
    }

    .st-key-active_main_tab .ai-nav-indicator::after {
        background: inherit;
        border-radius: inherit;
        content: "";
        filter: blur(7px);
        inset: -5px -9px;
        opacity: 0.58;
        position: absolute;
    }

    @keyframes aiNavGlowBreath {
        0%, 100% {
            opacity: 0.40;
            transform: translateY(-50%) scale(0.92);
        }
        50% {
            opacity: 0.85;
            transform: translateY(-50%) scale(1.08);
        }
    }

    @media (min-width: 768px) {
        .st-key-active_main_tab div[role="radiogroup"] {
            flex-direction: row !important;
            justify-content: center !important;
        }
        .st-key-active_main_tab div[role="radiogroup"] > label {
            flex: 0 0 auto;
        }
    }

    @media (min-width: 1024px) {
        .st-key-active_main_tab div[role="radiogroup"] > label {
            padding: 8px 20px !important;
        }
    }

    @media (max-width: 767px) {
        .st-key-active_main_tab div[role="radiogroup"] {
            align-items: center;
            flex-direction: column !important;
            gap: 12px;
            padding-bottom: 16px;
        }
        .st-key-active_main_tab div[role="radiogroup"] > label {
            font-size: 16px !important; /* Prevent zoom on iOS */
            min-height: auto;
            padding: 12px 16px !important;
            text-align: center;
            width: auto !important;
        }
        .st-key-active_main_tab div[role="radiogroup"] > label::before {
            left: 10px;
        }
    }

    </style>
    <style>
    @media (max-width: 767px) {
        .stButton > button {
            min-height: 48px !important;
            padding: 12px 16px !important;
            font-size: 16px !important;
        }
        
        /* Larger file cards for touch */
        [data-testid="stSidebar"] [class*="st-key-select_file_"] button {
            min-height: 80px !important;
            padding: 16px 12px !important;
        }
        
        /* Sidebar improvements for mobile */
        [data-testid="stSidebar"] {
            width: 280px !important;
        }
    }

    /* Prevent horizontal scroll */
    body {
        overflow-x: hidden;
    }

    /* Responsive text sizes */
    .login-heading {
        font-size: clamp(1.8rem, 5vw, 3.4rem) !important;
        color: #F8FAFC !important;
        font-weight: 700 !important;
        margin-bottom: 12px !important;
    }

    .login-subheading {
        font-size: clamp(0.9rem, 2.5vw, 1.1rem) !important;
        color: rgba(248, 250, 252, 0.85) !important;
        font-weight: 500 !important;
        margin-bottom: 24px !important;
    }

    .login-tagline {
        font-size: clamp(0.9rem, 2.5vw, 1rem);
    }

    /* Content responsiveness */
    .app-card {
        padding: clamp(1rem, 2vw, 1.5rem);
    }

    .metric-card {
        padding: clamp(1rem, 2vw, 1.5rem);
    }

    /* Table responsiveness */
    .scrollable {
        overflow-x: auto;
    }

    /* Ensure images are responsive */
    img {
        max-width: 100%;
        height: auto;
    }

    /* Form inputs on mobile */
    @media (max-width: 767px) {
        .glass-card .stTextInput input {
            font-size: 16px !important; /* Prevent zoom */
            padding: 16px !important;
            min-height: 48px !important;
        }
        
        .glass-card .stButton > button {
            min-height: 48px !important;
            font-size: 16px !important;
        }
    }

    /* Dashboard grid responsiveness */
    @media (max-width: 767px) {
        .dashboard-grid {
            grid-template-columns: 1fr;
            gap: 1rem;
        }
    }

    /* Dashboard grid responsiveness */
    @media (max-width: 767px) {
        .dashboard-grid {
            grid-template-columns: 1fr;
            gap: 1rem;
        }
    }

    /* Chat and other content */
    @media (max-width: 767px) {
        .stMarkdown {
            font-size: 14px;
        }
        
        .stDataFrame {
            font-size: 12px;
        }
    }

    /* ============================================ */
    /* ENHANCED RESPONSIVE DESIGN FOR ALL DEVICES */
    /* ============================================ */
    
    /* Ultra-responsive design for all screen sizes */
    @media (max-width: 320px) {
        /* Smartphones - extra small */
        .stApp {
            padding: 0.25rem;
        }
        .stSidebar {
            width: 100% !important;
        }
        h1 {
            font-size: 1.2rem !important;
        }
        h2 {
            font-size: 1rem !important;
        }
        .stButton > button {
            padding: 0.5rem !important;
            font-size: 0.75rem !important;
            min-width: 40px !important;
        }
    }

    @media (min-width: 321px) and (max-width: 480px) {
        /* Smartphones */
        .block-container {
            padding-left: 0.5rem !important;
            padding-right: 0.5rem !important;
        }
        .stColumn {
            padding: 0.5rem 0.25rem;
        }
        .metric-card {
            padding: 0.75rem !important;
        }
        .stMetric {
            font-size: 0.9rem;
        }
    }

    @media (min-width: 481px) and (max-width: 768px) {
        /* Tablets (portrait) */
        .block-container {
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }
        .dashboard-grid {
            grid-template-columns: repeat(2, 1fr);
        }
    }

    @media (min-width: 769px) and (max-width: 1024px) {
        /* Tablets (landscape) & small laptops */
        .dashboard-grid {
            grid-template-columns: repeat(3, 1fr);
        }
    }

    @media (min-width: 1025px) and (max-width: 1440px) {
        /* Laptops & desktops */
        .dashboard-grid {
            grid-template-columns: repeat(4, 1fr);
        }
    }

    @media (min-width: 1441px) {
        /* Large screens & projectors */
        .dashboard-grid {
            grid-template-columns: repeat(5, 1fr);
        }
        .block-container {
            max-width: 1920px;
        }
    }

    /* Projector & presentation mode */
    @media screen and (min-height: 1080px) {
        .stApp {
            background: #ffffff;
            font-size: 18px;
        }
        .stMarkdown {
            font-size: 18px;
        }
        h1 { font-size: 2.5rem; }
        h2 { font-size: 2rem; }
        h3 { font-size: 1.5rem; }
        .stButton > button {
            min-height: 60px;
            font-size: 18px;
            padding: 1rem;
        }
    }

    /* Touch-friendly interface (mobile & tablet) */
    @media (hover: none) {
        .stButton > button,
        [role="button"] {
            min-height: 48px;
            padding: 12px 16px;
            font-size: 16px;
        }
        input, select, textarea {
            min-height: 44px;
            padding: 12px;
        }
    }

    /* High DPI screens (retina) */
    @media (-webkit-min-device-pixel-ratio: 2), (min-resolution: 192dpi) {
        body {
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }
    }

    /* Landscape mode adjustments */
    @media (orientation: landscape) and (max-height: 500px) {
        .stApp {
            margin: 0;
            padding: 0;
        }
        .block-container {
            padding: 0.5rem;
        }
    }

    /* Print friendly */
    @media print {
        .stButton, [role="button"], .stSidebar,
        [data-testid="stHeader"],
        [data-testid="stToolbar"],
        footer {
            display: none !important;
        }
        body {
            background: white;
            color: black;
        }
        img {
            max-width: 100%;
            page-break-inside: avoid;
        }
    }

    /* Dark mode optimization */
    @media (prefers-color-scheme: dark) {
        :root {
            --background: #1a1a1a;
            --surface: #2d2d2d;
            --text: #e0e0e0;
        }
    }

    @keyframes selectedPillGlow {
        0%, 100% {
            box-shadow: 0 0 0 3px rgba(96, 165, 250, 0.18), 0 0 14px rgba(59, 130, 246, 0.28);
        }
        50% {
            box-shadow: 0 0 0 4px rgba(96, 165, 250, 0.24), 0 0 26px rgba(59, 130, 246, 0.48);
        }
    }

    .st-key-active_main_tab div[role="radiogroup"] > label[data-checked="true"] {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        color: #0f172a !important;
        font-weight: 800 !important;
        transform: translateZ(0);
        animation: none;
    }

    .st-key-active_main_tab div[role="radiogroup"] > label[data-checked="true"]:hover,
    .st-key-active_main_tab div[role="radiogroup"] > label.ai-nav-active:hover {
        transform: scale(1.035);
    }

    .st-key-dashboard_chart_type div[role="radiogroup"] > label[data-checked="true"] p,
    .st-key-dashboard_bar_orientation div[role="radiogroup"] > label[data-checked="true"] p {
        color: #0f172a !important;
        font-weight: 800 !important;
    }

    .st-key-dashboard_chart_type div[role="radiogroup"],
    .st-key-dashboard_bar_orientation div[role="radiogroup"] {
        justify-content: flex-start !important;
        margin-bottom: 0.5rem;
    }

    .st-key-dashboard_chart_type div[role="radiogroup"] > label,
    .st-key-dashboard_bar_orientation div[role="radiogroup"] > label {
        min-width: 132px !important;
        border-radius: 14px !important;
        border: 1px solid #d7e3f4 !important;
        background: #f8fbff !important;
    }

    .st-key-dashboard_chart_type div[role="radiogroup"] > label[data-checked="true"],
    .st-key-dashboard_bar_orientation div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #ecfeff 0%, #dbeafe 100%) !important;
        border: 2px solid #38bdf8 !important;
        color: #0f172a !important;
        animation: selectedPillGlow 2.2s ease-in-out infinite;
    }

    /* Reduced motion for accessibility */
    @media (prefers-reduced-motion: reduce) {
        *,
        *::before,
        *::after {
            animation-duration: 0.01ms !important;
            animation-iteration-count: 1 !important;
            transition-duration: 0.01ms !important;
        }
    }
    </style>
""", unsafe_allow_html=True)


st.markdown(
    """
    <style>
    /* Header sizing fixes */
    [data-testid="stHorizontalBlock"]:has(.st-key-header_brain_icon) {
        align-items: center !important;
        gap: 0.75rem !important;
    }
    .st-key-header_brain_icon button {
        width: 46px !important;
        min-width: 46px !important;
        height: 46px !important;
        min-height: 46px !important;
        padding: 0 !important;
        font-size: 1.55rem !important;
        line-height: 1 !important;
        border-radius: 12px !important;
    }
    .st-key-main_logout_btn button {
        min-height: 42px !important;
        padding: 0.55rem 0.75rem !important;
        white-space: nowrap !important;
    }

    /* Mobile layout repair: keep sidebar and main content from overlapping. */
    @media (max-width: 767px) {
        html, body, .stApp {
            width: 100% !important;
            max-width: 100% !important;
            overflow-x: hidden !important;
        }
        div[data-testid="stAppViewContainer"] {
            width: 100% !important;
            max-width: 100% !important;
            overflow-x: hidden !important;
        }
        .block-container,
        .main .block-container,
        section.main .block-container,
        div[data-testid="stMain"] .block-container {
            width: 100% !important;
            max-width: 100% !important;
            padding: 0.75rem !important;
            margin: 0 !important;
        }
        [data-testid="stSidebar"],
        .stSidebar {
            width: 100% !important;
            min-width: 0 !important;
            max-width: 100% !important;
            position: relative !important;
            transform: none !important;
            overflow-x: hidden !important;
        }
        [data-testid="stSidebar"] > div {
            width: 100% !important;
            max-width: 100% !important;
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
        }
        [data-testid="stHorizontalBlock"] {
            flex-wrap: wrap !important;
            gap: 0.5rem !important;
        }
        [data-testid="stHorizontalBlock"] > div {
            min-width: 0 !important;
        }
        [data-testid="stHorizontalBlock"]:has(.st-key-reset_chat_selection),
        [data-testid="stHorizontalBlock"]:has(.st-key-reset_dashboard_selection),
        [data-testid="stHorizontalBlock"]:has(.st-key-reset_compare_selection),
        [data-testid="stHorizontalBlock"]:has(.st-key-reset_capl_selection) {
            display: grid !important;
            grid-template-columns: minmax(0, 1fr) auto !important;
            align-items: center !important;
            column-gap: 0.75rem !important;
        }
        [data-testid="stHorizontalBlock"]:has(.st-key-reset_chat_selection) > div,
        [data-testid="stHorizontalBlock"]:has(.st-key-reset_dashboard_selection) > div,
        [data-testid="stHorizontalBlock"]:has(.st-key-reset_compare_selection) > div,
        [data-testid="stHorizontalBlock"]:has(.st-key-reset_capl_selection) > div {
            width: auto !important;
            min-width: 0 !important;
        }
        .st-key-reset_chat_selection button,
        .st-key-reset_dashboard_selection button,
        .st-key-reset_compare_selection button,
        .st-key-reset_capl_selection button {
            width: auto !important;
            min-width: 92px !important;
            white-space: nowrap !important;
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
        }
        [data-testid="stHorizontalBlock"]:has(.st-key-header_brain_icon) {
            display: grid !important;
            grid-template-columns: 52px minmax(0, 1fr) !important;
        }
        [data-testid="stHorizontalBlock"]:has(.st-key-main_logout_btn) {
            display: flex !important;
            align-items: center !important;
        }
        .st-key-main_logout_btn {
            margin-left: auto !important;
        }
        div[role="radiogroup"] {
            flex-direction: row !important;
            flex-wrap: wrap !important;
            justify-content: stretch !important;
            gap: 8px !important;
        }
        div[role="radiogroup"] > label {
            flex: 1 1 calc(50% - 8px) !important;
            min-width: 135px !important;
            width: auto !important;
        }
        .stSelectbox,
        .stMultiSelect,
        .stTextInput,
        .stTextArea,
        .stNumberInput,
        .stSlider,
        .stCheckbox,
        .stRadio,
        .stDownloadButton,
        .stFileUploader,
        .stDataFrame,
        [data-testid="stDataFrame"],
        [data-testid="stMetric"],
        [data-testid="stPlotlyChart"],
        iframe {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
        }
        .stDataFrame,
        [data-testid="stDataFrame"],
        [data-testid="stTable"],
        [data-testid="stPlotlyChart"] {
            overflow-x: auto !important;
        }
        [data-testid="column"],
        [data-testid="stColumn"] {
            min-width: min(100%, 260px) !important;
        }
        [data-testid="stMetric"] {
            overflow-wrap: anywhere !important;
        }
        .app-card,
        .file-chip-wrap,
        .file-chip {
            max-width: 100% !important;
            overflow-wrap: anywhere !important;
        }
        .dashboard-grid {
            grid-template-columns: 1fr !important;
        }
    }

    @media (min-width: 768px) and (max-width: 1180px) {
        [data-testid="stHorizontalBlock"] {
            gap: 0.75rem !important;
        }
        [data-testid="column"],
        [data-testid="stColumn"] {
            min-width: 0 !important;
        }
        .stDataFrame,
        [data-testid="stDataFrame"],
        [data-testid="stTable"],
        [data-testid="stPlotlyChart"],
        iframe {
            max-width: 100% !important;
            overflow-x: auto !important;
        }
        div[role="radiogroup"] {
            flex-wrap: wrap !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <style>
    .ai-os-shell {
        margin: 10px 0 14px;
        padding: 16px 18px;
        border: 1px solid rgba(148, 163, 184, 0.24);
        border-radius: 18px;
        background: rgba(248, 251, 255, 0.72);
        box-shadow: 0 18px 50px rgba(15, 23, 42, 0.08);
        backdrop-filter: blur(14px);
        animation: aiOsRise 0.45s ease-out both;
    }
    .ai-os-kicker {
        color: #2563eb;
        font-size: 0.78rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }
    .ai-os-title {
        margin-top: 4px;
        color: #0f172a;
        font-size: 1.08rem;
        font-weight: 800;
    }
    .ai-os-loop,
    .ai-os-metrics {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-top: 10px;
    }
    .ai-os-loop span,
    .ai-os-metrics span {
        padding: 6px 10px;
        border: 1px solid rgba(147, 197, 253, 0.45);
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.62);
        color: #173152;
        font-size: 0.84rem;
        transition: transform 0.18s ease, background-color 0.18s ease;
    }
    .ai-os-loop span:hover,
    .ai-os-metrics span:hover {
        transform: translateY(-1px) scale(1.01);
        background: rgba(239, 246, 255, 0.92);
    }
    @keyframes aiOsRise {
        from { opacity: 0; transform: translateY(8px); }
        to { opacity: 1; transform: translateY(0); }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# -------------------------------
# MAIN TAB NAVIGATION
# -------------------------------
# Creates the horizontal tab navigation with custom styling.
# Each tab corresponds to a major feature area of the application.




main_tab_options = TAB_OPTIONS

# ==============================
# ROUTER / STATE FIREWALL INITIALIZATION
# The router is the single source of truth for active tabs. Automatic tab
# switching is disabled unless auto_tab_switch_enabled is explicitly set.
# ==============================
init_state_firewall()
init_tab_memory()
init_router(main_tab_options[0])
ensure_context_memory()
if st.session_state.get("auto_tab_switch_enabled", False):
    apply_auto_tab_suggestion(main_tab_options)

# ==============================
# ANIMATED TAB COLOR SYSTEM
# Colors are generated once, stored in session_state.tab_colors, and reused.
# ==============================
tab_colors = ensure_tab_glow_colors(main_tab_options)
tab_color_css = "\n".join(
    (
        ".st-key-active_main_tab div[role=\"radiogroup\"] > label:nth-of-type({index}) {{ "
        "--tab-glow: {color}; --tab-glow-rgb: {red}, {green}, {blue}; }}"
    ).format(
        index=index,
        color=tab_colors[tab_name],
        red=hex_to_rgb_values(tab_colors[tab_name])[0],
        green=hex_to_rgb_values(tab_colors[tab_name])[1],
        blue=hex_to_rgb_values(tab_colors[tab_name])[2],
    )
    for index, tab_name in enumerate(main_tab_options, start=1)
)

st.markdown(
    f"""
    <style>
    {tab_color_css}

    .st-key-active_main_tab {{
        --tab-neutral-bg: rgba(248, 250, 252, 0.76);
        --tab-neutral-border: rgba(148, 163, 184, 0.24);
        margin-top: 10px !important;
        margin-bottom: 10px !important;
    }}

    .st-key-active_main_tab .ai-nav-indicator {{
        display: none !important;
    }}

    .st-key-active_main_tab div[role="radiogroup"] {{
        align-items: stretch !important;
        border-bottom: none !important;
        display: flex !important;
        gap: clamp(10px, 1.6vw, 16px) !important;
        justify-content: center !important;
        overflow: visible !important;
        padding: 5px 4px 14px !important;
    }}

    .st-key-active_main_tab div[role="radiogroup"] > label > div:first-child {{
        display: none !important;
    }}

    .st-key-active_main_tab div[role="radiogroup"] > label {{
        isolation: isolate !important;
        overflow: hidden !important;
        position: relative !important;
        background: var(--tab-neutral-bg) !important;
        border: 1px solid var(--tab-neutral-border) !important;
        border-left: 5px solid transparent !important;
        border-radius: 12px !important;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06) !important;
        color: #64748b !important;
        cursor: pointer !important;
        opacity: 0.9;
        padding: 11px 20px 11px 16px !important;
        transform: translate3d(0, 0, 0) scale(1);
        transition:
            background 260ms ease-out,
            border-color 260ms ease-out,
            box-shadow 260ms ease-out,
            color 260ms ease-out,
            opacity 260ms ease-out,
            transform 260ms cubic-bezier(0.16, 1, 0.3, 1) !important;
    }}

    .st-key-active_main_tab div[role="radiogroup"] > label::before {{
        animation: activeTabIndicatorIn 280ms ease-out both;
        background: var(--tab-glow, #38bdf8) !important;
        border-radius: 999px !important;
        bottom: 18% !important;
        box-shadow: none !important;
        content: "" !important;
        filter: none !important;
        left: -5px !important;
        opacity: 0 !important;
        position: absolute !important;
        top: 18% !important;
        transform: scaleY(0.52) !important;
        transition: opacity 260ms ease-out, transform 260ms ease-out !important;
        width: 5px !important;
        z-index: 2 !important;
    }}

    .st-key-active_main_tab div[role="radiogroup"] > label::after {{
        animation: activeTabLightSweep 4.2s linear infinite;
        background:
            linear-gradient(115deg,
                transparent 0%,
                transparent 33%,
                rgba(255, 255, 255, 0.14) 42%,
                rgba(255, 255, 255, 0.72) 50%,
                rgba(255, 255, 255, 0.16) 58%,
                transparent 68%,
                transparent 100%) !important;
        content: "" !important;
        display: block !important;
        inset: -45% -70% !important;
        opacity: 0 !important;
        pointer-events: none !important;
        position: absolute !important;
        transform: translateX(-68%) rotate(0.001deg);
        z-index: 0 !important;
    }}

    .st-key-active_main_tab div[role="radiogroup"] > label:hover {{
        background: rgba(255, 255, 255, 0.94) !important;
        border-color: rgba(var(--tab-glow-rgb, 56, 189, 248), 0.34) !important;
        box-shadow:
            0 10px 24px rgba(15, 23, 42, 0.08),
            0 0 18px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.18) !important;
        color: #1e293b !important;
        opacity: 1;
        transform: translate3d(3px, -2px, 0) scale(1.045) !important;
    }}

    .st-key-active_main_tab div[role="radiogroup"] > label:nth-of-type(even):hover {{
        transform: translate3d(-3px, -2px, 0) scale(1.045) !important;
    }}

    .st-key-active_main_tab div[role="radiogroup"] > label[data-checked="true"],
    .st-key-active_main_tab div[role="radiogroup"] > label.ai-nav-active {{
        animation:
            activeTabGlowIn 320ms ease-out both,
            activeTabNeonBreath 4.8s ease-in-out infinite,
            activeTabGradientDrift 5.2s linear infinite !important;
        background:
            linear-gradient(110deg,
                rgba(var(--tab-glow-rgb, 56, 189, 248), 0.30) 0%,
                rgba(var(--tab-glow-rgb, 56, 189, 248), 0.13) 34%,
                rgba(255, 255, 255, 0.82) 50%,
                rgba(var(--tab-glow-rgb, 56, 189, 248), 0.18) 66%,
                rgba(var(--tab-glow-rgb, 56, 189, 248), 0.32) 100%) !important;
        background-size: 240% 100% !important;
        border-color: rgba(var(--tab-glow-rgb, 56, 189, 248), 0.72) !important;
        border-left-color: var(--tab-glow, #38bdf8) !important;
        box-shadow:
            0 0 0 1px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.26),
            0 0 20px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.50),
            0 0 44px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.24),
            inset 0 0 24px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.16) !important;
        color: #0f172a !important;
        opacity: 1;
        transform: translate3d(0, -1px, 0) scale(1.035) !important;
    }}

    .st-key-active_main_tab div[role="radiogroup"] > label[data-checked="true"]::before,
    .st-key-active_main_tab div[role="radiogroup"] > label.ai-nav-active::before {{
        animation:
            activeTabIndicatorIn 280ms ease-out both,
            activeTabIndicatorPulse 3.2s ease-in-out infinite !important;
        box-shadow:
            0 0 11px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.88),
            0 0 24px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.46) !important;
        opacity: 1 !important;
        transform: scaleY(1) !important;
    }}

    .st-key-active_main_tab div[role="radiogroup"] > label[data-checked="true"]::after,
    .st-key-active_main_tab div[role="radiogroup"] > label.ai-nav-active::after {{
        opacity: 1 !important;
    }}

    .st-key-active_main_tab div[role="radiogroup"] > label p {{
        color: inherit !important;
        font-weight: 800 !important;
        position: relative !important;
        z-index: 1 !important;
    }}

    @keyframes activeTabGlowIn {{
        from {{
            opacity: 0.68;
            transform: translate3d(0, 2px, 0) scale(0.985);
        }}
        to {{
            opacity: 1;
            transform: translate3d(0, -1px, 0) scale(1.035);
        }}
    }}

    @keyframes activeTabGradientDrift {{
        0% {{ background-position: 0% 50%; }}
        100% {{ background-position: 240% 50%; }}
    }}

    @keyframes activeTabLightSweep {{
        0% {{ transform: translateX(-68%) skewX(-16deg); }}
        46% {{ transform: translateX(68%) skewX(-16deg); }}
        100% {{ transform: translateX(68%) skewX(-16deg); }}
    }}

    @keyframes activeTabNeonBreath {{
        0%, 100% {{
            box-shadow:
                0 0 0 1px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.22),
                0 0 16px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.38),
                0 0 34px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.18),
                inset 0 0 18px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.12);
        }}
        50% {{
            box-shadow:
                0 0 0 1px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.36),
                0 0 26px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.62),
                0 0 56px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.32),
                inset 0 0 26px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.22);
        }}
    }}

    @keyframes activeTabIndicatorIn {{
        from {{ opacity: 0; transform: scaleY(0.35); }}
        to {{ opacity: 1; transform: scaleY(1); }}
    }}

    @keyframes activeTabIndicatorPulse {{
        0%, 100% {{
            top: 20%;
            bottom: 20%;
            box-shadow:
                0 0 8px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.66),
                0 0 18px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.34);
        }}
        50% {{
            top: 12%;
            bottom: 12%;
            box-shadow:
                0 0 14px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.96),
                0 0 30px rgba(var(--tab-glow-rgb, 56, 189, 248), 0.54);
        }}
    }}

    @media (max-width: 767px) {{
        .st-key-active_main_tab div[role="radiogroup"] {{
            flex-direction: column !important;
            gap: 10px !important;
        }}
        .st-key-active_main_tab div[role="radiogroup"] > label {{
            width: 100% !important;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)
active_main_tab = render_tab_router("Open Section")

# -------------------------------
# ==============================
# TAB ROUTING
# The visible active-tab radio remains identical; only the tab body execution
# moved into dedicated modules for lazy execution and faster switching.
# ==============================
fn.query_params = query_params
if active_main_tab == main_tab_options[0]:
    with tab_state_scope("chat"):
        render_chat_tab()
elif active_main_tab == main_tab_options[1]:
    with tab_state_scope("dashboard"):
        render_dashboard_tab()
elif active_main_tab == main_tab_options[2]:
    with tab_state_scope("compare"):
        render_compare_tab()
elif active_main_tab == main_tab_options[3]:
    with tab_state_scope("capl"):
        render_capl_tab()
