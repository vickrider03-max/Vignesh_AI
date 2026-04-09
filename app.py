import html, re, hashlib, os, json, base64, pickle
import uuid
import urllib.parse
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from io import BytesIO
from pytz import timezone
import time

import docx, openpyxl, pdfplumber, streamlit as st
from docx.text.paragraph import Paragraph
from docx.table import Table
import pandas as pd
from openpyxl.styles import PatternFill
from pptx import Presentation
from bs4 import BeautifulSoup
from PIL import Image, ImageDraw, ImageFont
import plotly.express as px
import plotly.graph_objects as go

from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.llms import HuggingFacePipeline
from langchain_community.vectorstores import FAISS
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.runnables import RunnablePassthrough
from langchain_text_splitters import RecursiveCharacterTextSplitter

PREVIEW_TOKENS = {}  # token -> {'file_name': str, 'timestamp': datetime}
PREVIEW_STORE = {}   # token -> file_dict

PREVIEW_DATA_FILE = "preview_data.pkl"

# Preview token helpers:
# Used by the document preview flow opened in a separate browser tab/window.
# These functions persist preview state so a selected file can still be rendered
# even after Streamlit reruns the main app script.
def load_preview_data():
    """Load preview data from file"""
    global PREVIEW_TOKENS, PREVIEW_STORE
    if os.path.exists(PREVIEW_DATA_FILE):
        try:
            with open(PREVIEW_DATA_FILE, "rb") as f:
                data = pickle.load(f)
                PREVIEW_TOKENS = data.get("tokens", {})
                PREVIEW_STORE = data.get("store", {})
        except Exception as e:
            st.warning(f"Could not load preview data: {e}")
            PREVIEW_TOKENS = {}
            PREVIEW_STORE = {}

def save_preview_data():
    """Save preview data to file"""
    try:
        data = {
            "tokens": PREVIEW_TOKENS,
            "store": PREVIEW_STORE
        }
        with open(PREVIEW_DATA_FILE, "wb") as f:
            pickle.dump(data, f)
    except Exception as e:
        st.warning(f"Could not save preview data: {e}")

def cleanup_expired_preview_tokens():
    """Remove preview tokens older than 1 hour to prevent memory accumulation."""
    now = datetime.now()
    expired_tokens = []
    for token, data in PREVIEW_TOKENS.items():
        if now - data['timestamp'] > timedelta(hours=1):
            expired_tokens.append(token)
    
    for token in expired_tokens:
        del PREVIEW_TOKENS[token]
        if token in PREVIEW_STORE:
            del PREVIEW_STORE[token]
    
    # Save updated data
    save_preview_data()


@st.cache_data(show_spinner=False)
def get_needle_minimalist_logo():
    try:
        import matplotlib.pyplot as plt
        import numpy as np
        from matplotlib.patches import Polygon
    except Exception:
        st.warning("Mercedes logo generation requires matplotlib and numpy. Install these packages to view the animated logo.")
        return None

    frames = []

    SILVER_GREY = '#A0A0A0'
    STAR_LIGHT = '#DCDCDC'
    STAR_SHADOW = '#B8B8B8'

    for angle_deg in range(360, 0, -15):
        fig, ax = plt.subplots(figsize=(4, 4), facecolor='none')
        ax.set_xlim(-1.1, 1.1)
        ax.set_ylim(-1.1, 1.1)
        ax.axis('off')

        raw_scale = np.cos(np.deg2rad(angle_deg))
        flip_scale = raw_scale if abs(raw_scale) > 0.08 else (0.08 * np.sign(raw_scale + 0.001))

        theta = np.linspace(0, 2 * np.pi, 150)
        ax.plot(np.cos(theta) * flip_scale, np.sin(theta),
                color=SILVER_GREY, linewidth=3.0, zorder=1)

        base_angles = np.deg2rad([90, 210, 330])
        for angle in base_angles:
            center = [0, 0]
            tip = [0.88 * np.cos(angle) * flip_scale, 0.88 * np.sin(angle)]
            side_l = [0.11 * np.cos(angle + 2.15) * flip_scale, 0.11 * np.sin(angle + 2.15)]
            side_r = [0.11 * np.cos(angle - 2.15) * flip_scale, 0.11 * np.sin(angle - 2.15)]

            if flip_scale > 0:
                c_l, c_r = STAR_LIGHT, STAR_SHADOW
            else:
                c_l, c_r = STAR_SHADOW, STAR_LIGHT

            ax.add_patch(Polygon([center, tip, side_l], facecolor=c_l, zorder=2))
            ax.add_patch(Polygon([center, tip, side_r], facecolor=c_r, zorder=2))

        buf = BytesIO()
        plt.savefig(buf, transparent=True, format='png', bbox_inches='tight', pad_inches=0)
        buf.seek(0)
        frames.append(Image.open(buf))
        plt.close(fig)

    gif_buf = BytesIO()
    if frames:
        frames[0].save(
            gif_buf,
            format='GIF',
            save_all=True,
            append_images=frames[1:],
            duration=60,
            loop=0,
            disposal=2
        )

    return base64.b64encode(gif_buf.getvalue()).decode('utf-8')


# -------------------------------
# STREAMLIT PAGE CONFIG
# -------------------------------
st.set_page_config(page_title="🧠 IntelliDoc AI– Smart Document Assistant", layout="wide")

# Load preview data from file
load_preview_data()

# Clean up expired preview tokens on app start
cleanup_expired_preview_tokens()

try:
    logo_data = get_needle_minimalist_logo()
except Exception:
    logo_data = None

st.markdown(
    """
    <style>
        :root {
            --brand: #1f4f91;
            --brand-soft: #eaf2ff;
            --border: #d7e3f4;
            --text-soft: #51627a;
            --panel: #f8fbff;
            --success-bg: #edf8f1;
            --warning-bg: #fff8e8;
        }
        .app-card {
            background: var(--panel);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 14px 16px;
            margin: 8px 0 14px 0;
        }
        .app-card h4 {
            margin: 0 0 6px 0;
            color: var(--brand);
            font-size: 15px;
        }
        .app-muted {
            color: var(--text-soft);
            font-size: 13px;
            margin: 0;
        }
        .dashboard-grid {
            display: flex;
            justify-content: space-between;
            gap: 20px;
            margin-bottom: 30px;
        }

        .metric-card {
            flex: 1;
            background: #ffffff;
            padding: 18px;
            border-radius: 12px;
            border: 1px solid #e2e5ed;
            box-shadow: 0 2px 6px rgba(30, 64, 175, 0.08);
            text-align: center;
            transition: all 0.25s ease;
        }

        .card-label {
            display: block;
            font-size: 0.75rem;
            font-weight: 600;
            color: #9ca3af;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
        }

        .card-value {
            display: block;
            font-size: 1.35rem;
            font-weight: 700;
            color: #1e40af;
        }

        .metric-card:hover {
            border-color: #1e40af;
            box-shadow: 0 6px 18px rgba(30, 64, 175, 0.12);
        }
            font-weight: 600;
            line-height: 1.1;
        }
        .file-chip-wrap {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-top: 8px;
        }
        .file-chip {
            background: var(--brand-soft);
            color: var(--brand);
            border: 1px solid var(--border);
            border-radius: 999px;
            display: inline-block;
            font-size: 12px;
            padding: 5px 10px;
        }
        .section-note {
            color: var(--text-soft);
            font-size: 13px;
            margin: 0 0 6px 0;
        }
        .brand-hero {
            display: flex;
            flex-direction: column;
            align-items: flex-start;
            justify-content: center;
            gap: 4px;
            width: 100%;
            margin: 0 0 8px 0;
            text-align: left;
        }
        .brand-hero img {
            width: 120px;
            max-width: 100%;
            height: auto;
            display: block;
        }
        .brand-title {
            color: #173152;
            font-size: 1.85rem;
            font-weight: 700;
            line-height: 1.1;
            margin: 0;
        }
        .brand-subtitle {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            color: #666;
            font-size: 0.68rem;
            margin: 0;
            font-weight: 500;
            letter-spacing: 0.18rem;
            text-transform: uppercase;
        }
        div[data-baseweb="select"] span[data-baseweb="tag"] {
            background: #dff3e4 !important;
            border: 1px solid #b7dec1 !important;
            border-radius: 999px !important;
            color: #2f5d3a !important;
        }
        div[data-baseweb="select"] span[data-baseweb="tag"] * {
            color: #2f5d3a !important;
        }
        div[data-baseweb="select"] span[data-baseweb="tag"] svg {
            fill: #4c7a57 !important;
        }
        div[data-baseweb="select"] span[data-baseweb="tag"]:hover {
            background: #d3edd9 !important;
            border-color: #a7d2b2 !important;
        }
    </style>
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
    ("chat_summary_downloads", {"images": [], "tables": []}),
    ("messages", []),
    ("welcome_shown", False),
]:
    if key not in st.session_state:
        st.session_state[key] = default_value


def render_status_strip():
    if not st.session_state.get("is_authenticated"):
        return

    if 'start_time' not in st.session_state or st.session_state.start_time is None:
        st.session_state.start_time = time.time()

    elapsed = int(time.time() - st.session_state.start_time)
    hours, rem = divmod(elapsed, 3600)
    mins, secs = divmod(rem, 60)
    timer_str = f"{hours:02d}:{mins:02d}:{secs:02d}"

    username = st.session_state.get("logged_in_username") or "Vignesh"
    role = st.session_state.get("user_role") or "User"
    available_files = len(st.session_state.get("selected_files", []))

    status_html = f"""
    <div class="dashboard-grid">
        <div class="metric-card">
            <span class="card-label">User</span>
            <span class="card-value">{html.escape(username)}</span>
        </div>
        <div class="metric-card">
            <span class="card-label">Role</span>
            <span class="card-value">{html.escape(str(role).title())}</span>
        </div>
        <div class="metric-card">
            <span class="card-label">Available Files</span>
            <span class="card-value">{available_files}</span>
        </div>
        <div class="metric-card">
            <span class="card-label">Usage Time</span>
            <span class="card-value">{timer_str}</span>
        </div>
    </div>
    """

    st.markdown(status_html, unsafe_allow_html=True)

col_logo, col_status = st.columns([2, 4])
with col_logo:
    if logo_data:
        st.markdown(
            f"""
            <div class="brand-hero">
                <img src="data:image/gif;base64,{logo_data}" alt="Mercedes-Benz logo" />
                <p class="brand-subtitle">Mercedes-Benz</p>
                <h1 class="brand-title">Vignesh_AI🧠 IntelliDoc AI– Smart Document Assistant</h1>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div class="brand-hero">
                <h1 class="brand-title">🧠 IntelliDoc AI– Smart Document Assistant</h1>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.info("Mercedes logo generation requires matplotlib and numpy. Install these packages to view the animated logo.")

with col_status:
    if st.session_state.is_authenticated:
        if not st.session_state.get('welcome_shown', False):
            st.toast(f"🎉 Welcome back, {st.session_state.logged_in_username}! We're thrilled to have you here.", icon="🎉")
            st.session_state.welcome_shown = True

        render_status_strip()

        creator_timestamp = None
        if st.session_state.user_role == "creator" and st.session_state.login_history:
            creator_entries = [
                entry for entry in st.session_state.login_history
                if entry.get("username") == st.session_state.logged_in_username and entry.get("role") == "creator"
            ]
            if creator_entries:
                creator_timestamp = creator_entries[-1].get("timestamp")

        status_message = f"Logged in as {st.session_state.logged_in_username} ({st.session_state.user_role})"
        if creator_timestamp:
            status_message += f"\nLogin time: {creator_timestamp}"

        st.markdown(f"**{status_message}**")
        logout_clicked = st.button("Logout", key="logout_button")
        if logout_clicked:
            active_file = "active_users.json"
            if os.path.exists(active_file):
                with open(active_file, "r") as f:
                    active_users = json.load(f)
                active_users = [u for u in active_users if u["username"] != st.session_state.logged_in_username]
                with open(active_file, "w") as f:
                    json.dump(active_users, f)
            st.session_state.is_authenticated = False
            st.session_state.logged_in_username = ""
            st.session_state.user_role = None
            st.session_state.user_session_start_time = None
            st.session_state.selected_files = []
            st.session_state.file_texts = {}
            st.session_state.vector_stores = {}
            st.session_state.chat_file_selection = []
            st.session_state.chat_summary_downloads = {"images": [], "tables": []}
            st.session_state.messages = []
            st.session_state.welcome_shown = False
            st.success("Logged out successfully.")
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
    st.session_state.chat_summary_downloads = {"images": [], "tables": []}
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

# Load README.txt from file system
def load_readme_text():
    """Load README from file if it exists, otherwise use default"""
    readme_path = "README.txt"
    if os.path.exists(readme_path):
        try:
            with open(readme_path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
        except Exception as e:
            st.warning(f"Could not load README.txt: {e}")
    # Return default README if file not found
    return """README.txt
==========

Multi-Utility File & CAPL Analyzer Tool
=======================================

Overview
--------
This Streamlit-based application helps manage, analyze, and compare files, with special support for CAPL scripts. It combines file dashboards, comparisons, CAPL analysis, and AI-assisted code fixing in a single platform.

---

App Layout & Workflow
---------------------

+----------------------+
| Sidebar              |
|----------------------|
| Upload Files         |
| Select Files         |
| Filter CAPL (.can)   |
+----------------------+
          |
          v
+------------------------------------------+
| Main Tabs:                               |
|------------------------------------------|
| [ Chat ] [ Dashboard ] [ Compare ] [ CAPL ] |
+------------------------------------------+

Tab Workflow:
-------------

1. Chat Tab
   +-----------------------------+
   | Ask questions about files   |
   | Semantic / AI answers       |
   +-----------------------------+
          |
          v
   (Uses uploaded files & AI backend)

2. Dashboard Tab
   +------------------------------+
   | Select a file from sidebar   |
   | Visualize trends (Excel/CSV) |
   | Aggregated stats             |
   +------------------------------+
          |
          v
   (Optional download of analyzed results)

3. Compare Tab
   +--------------------------------+
   | Multi-file selection           |
   | Inline word-level differences  |
   | Download comparison Excel      |
   +--------------------------------+
          |
          v
   (At least 2 files required)

4. CAPL Tab
   +-------------------------------------------+
   | Select existing CAPL file or create new   |
   | Compile & analyze code                    |
   | View issues / suggestions                 |
   | AI-assisted fix / Apply fix / Save file   |
   +-------------------------------------------+
          |
          v
   (Updates session state & file texts)

---

Feature Summary
---------------

1. Chat / RAG Interface
   - Ask questions about selected files.
   - Context-aware AI responses.

2. File Dashboard
   - Test report analysis and visualization.
   - Downloadable Excel summaries.

3. Compare Files
   - Multi-file comparison.
   - Inline word-level differences.
   - Downloadable Excel comparison.

4. CAPL Compiler & Analyzer
   - Upload or create CAPL scripts (.can/.txt).
   - Syntax highlighting & code analysis.
   - AI-assisted code fixes.
   - Save new or corrected CAPL files.

5. Interactive UI
   - Tabs for workflows.
   - Reset buttons for selections and results.
   - Expandable live editor for CAPL scripts.

6. AI Integration
   - Auto-correct CAPL code.
   - Chat-based file analysis.

7. Session Management
   - Tracks uploaded files and selected files per tab.
   - Maintains last analyzed CAPL file and issues.

---

How to Use
----------

1. Setup
   - Python >= 3.10
   - Install dependencies:
     pip install streamlit openai pandas plotly
   - Configure AI backend / API keys if using AI features.

2. Run
   - streamlit run app.py

3. Sidebar
   - Upload files.
   - Select files to be available in tabs.
   - Optionally filter CAPL scripts.

4. Tabs
   - Chat: Ask questions about uploaded files.
   - Dashboard: Visualize file content, trends, and statistics.
   - Compare: Choose 2+ files and see word-level differences.
   - CAPL: Edit, compile, analyze CAPL scripts, AI fixes, save.

5. CAPL AI Fix
   - Click "Suggest Fix" -> review AI suggestion -> click "Use Suggested Fix".

6. Reset Buttons
   - Clear selections and results in each tab.

---

Tips & Notes
------------

- Limit comparisons to <2000 lines for performance.
- CAPL files must end with .can.
- AI features require backend availability.
- Use "Include all .txt files as CAPL" cautiously.

---

ASCII Workflow Example
----------------------

Sidebar:
--------
+----------------------+
| Upload Files         |
| Select Files         |
| Filter CAPL (.can)   |
+----------------------+

Tabs:
-----
[ Chat ] -> AI Q&A using uploaded files
[ Dashboard ] -> File stats & trends -> Excel download
[ Compare ] -> Multi-file diff -> HTML + Excel
[ CAPL ] -> Edit/Compile/Analyze -> AI Fix -> Save

CAPL AI Flow:
-------------
Create/Edit CAPL
    |
    v
Compile & Analyze
    |
    v
AI Suggest Fix? -- Yes --> Review & Apply --> Update Editor
    |
    No
    v
Save CAPL Script

---

Credits
-------
- Built with Streamlit, Pandas, Plotly, OpenAI API
- CAPL analysis inspired by automotive testing
- AI auto-fix powered by LLMs

---

Contact
-------
For support or feedback, contact vigneshs075@gmail.com.
"""

README_TEXT = load_readme_text()

# -------------------------------
# DOCUMENT PREVIEW FUNCTION
# -------------------------------
# Preview processing helpers:
# These functions support the standalone preview page and also provide extracted
# text/data reused by Chat, Dashboard, Compare, and CAPL when files are selected.
def ensure_file_processed(file_name):
    file_info = get_uploaded_file_entry(file_name)
    if not file_info:
        return
    file_name_lower = file_name.lower()

    if file_name not in st.session_state.file_texts:
        st.session_state.file_texts[file_name] = extract_text(file_name, file_info["bytes"])

    if file_name_lower.endswith(".xlsx") and file_name not in st.session_state.excel_data_by_file:
        st.session_state.excel_data_by_file[file_name] = extract_excel_data(file_name, file_info["bytes"])


def extract_text(file_name, file_bytes):
    """Extract comprehensive text content from various file formats including tables and metadata."""
    text_parts = []
    bio = BytesIO(file_bytes)
    file_name_lower = file_name.lower()
    
    try:
        if file_name_lower.endswith(".pdf"):
            text_parts.extend(extract_pdf_content(bio))
        elif file_name_lower.endswith(".docx"):
            text_parts.extend(extract_docx_content(bio))
        elif file_name_lower.endswith(".pptx"):
            text_parts.extend(extract_pptx_content(bio))
        elif file_name_lower.endswith(".xlsx"):
            text_parts.extend(extract_xlsx_content(bio))
        elif file_name_lower.endswith((".html", ".htm")):
            text_parts.extend(extract_html_content(bio))
        elif file_name_lower.endswith(".txt"):
            text_parts.append(("TEXT", bio.read().decode("utf-8", errors="ignore")))
        elif file_name_lower.endswith(".can"):
            text_parts.append(("TEXT", bio.read().decode("utf-8", errors="ignore")))
        else:
            text_parts.append(("UNSUPPORTED", f"Unsupported file format: {file_name_lower}"))
    
    except Exception as e:
        text_parts.append(("ERROR", f"Error extracting content: {str(e)}"))
    
    # Combine all extracted content
    combined_text = ""
    for content_type, content in text_parts:
        if content_type == "TEXT":
            combined_text += content + "\n"
        elif content_type == "TABLE":
            combined_text += f"\nTABLE:\n{content}\n"
        elif content_type == "IMAGE":
            combined_text += f"\n[IMAGE: {content}]\n"
        elif content_type == "EMBEDDED_IMAGE":
            combined_text += f"\n[EMBEDDED_IMAGE: {content}]\n"
        elif content_type == "METADATA":
            combined_text += f"\n{content}\n"
        elif content_type == "ERROR":
            combined_text += f"\nERROR: {content}\n"
        elif content_type == "UNSUPPORTED":
            combined_text += f"\n{content}\n"
    
    return combined_text.strip()


def extract_pdf_content(bio):
    """Extract text, tables, and metadata from PDF."""
    content = []
    try:
        with pdfplumber.open(bio) as pdf:
            # Add metadata
            if pdf.metadata:
                metadata = []
                for key, value in pdf.metadata.items():
                    if value:
                        metadata.append(f"{key}: {value}")
                if metadata:
                    content.append(("METADATA", "PDF Metadata:\n" + "\n".join(metadata)))
            
            content.append(("METADATA", f"Total Pages: {len(pdf.pages)}"))
            
            # Extract content from each page
            for i, page in enumerate(pdf.pages):
                page_content = []
                
                # Extract text
                page_text = page.extract_text() or ""
                if page_text.strip():
                    page_content.append(f"Page {i+1} Text:\n{page_text}")
                
                # Extract tables
                tables = page.extract_tables()
                if tables:
                    for j, table in enumerate(tables):
                        if table and any(any(cell for cell in row) for row in table):
                            table_text = "\n".join([" | ".join(str(cell) if cell else "" for cell in row) for row in table])
                            page_content.append(f"Page {i+1} Table {j+1}:\n{table_text}")
                
                # Check for images (basic detection)
                if hasattr(page, 'images') and page.images:
                    content.append(("IMAGE", f"Page {i+1}: {len(page.images)} images detected"))
                
                if page_content:
                    content.append(("TEXT", "\n\n".join(page_content)))
    
    except Exception as e:
        content.append(("ERROR", f"PDF extraction failed: {str(e)}"))
    
    return content


def table_to_png_bytes(table_data, title=None):
    """Render table rows as a PNG image and return the bytes."""
    try:
        font = ImageFont.load_default()
    except Exception:
        font = None

    padding = 10
    row_height = 22
    col_padding = 18

    # Normalize table data
    normalized_table = [[str(cell) for cell in row] for row in table_data]
    col_widths = []
    for col_idx in range(len(normalized_table[0])):
        col_width = max(len(row[col_idx]) for row in normalized_table) * 7 + col_padding
        col_widths.append(col_width)

    width = sum(col_widths) + padding * 2
    height = row_height * len(normalized_table) + padding * 2
    if title:
        height += row_height

    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    y = padding

    if title:
        draw.text((padding, y), title, fill="black", font=font)
        y += row_height

    for row in normalized_table:
        x = padding
        for col_idx, cell in enumerate(row):
            draw.text((x, y), cell, fill="black", font=font)
            x += col_widths[col_idx]
        y += row_height

    output = BytesIO()
    image.save(output, format="PNG")
    return output.getvalue()


def image_bytes_to_png_bytes(image_bytes):
    """Convert an uploaded image to PNG bytes."""
    with Image.open(BytesIO(image_bytes)) as image:
        png_buffer = BytesIO()
        image.save(png_buffer, format="PNG")
        return png_buffer.getvalue()


def dataframe_to_table_rows(df):
    """Convert a dataframe to table rows suitable for PNG rendering."""
    safe_df = df.fillna("")
    rows = [list(map(str, safe_df.columns.tolist()))]
    rows.extend([list(map(str, row)) for row in safe_df.values.tolist()])
    return rows


def crop_pdf_region_to_png(page, bbox, resolution=150):
    """Crop a rectangular region from a PDF page and return it as PNG bytes."""
    cropped_page = page.crop(bbox)
    cropped_image = cropped_page.to_image(resolution=resolution)
    output = BytesIO()
    cropped_image.original.save(output, format="PNG")
    return output.getvalue()


def extract_docx_content(bio):
    """Extract text, tables, and images from DOCX."""
    content = []
    try:
        doc = docx.Document(bio)
        
        # Extract metadata
        core_props = doc.core_properties
        metadata = []
        if core_props.title:
            metadata.append(f"Title: {core_props.title}")
        if core_props.author:
            metadata.append(f"Author: {core_props.author}")
        if core_props.created:
            metadata.append(f"Created: {core_props.created}")
        if metadata:
            content.append(("METADATA", "Document Metadata:\n" + "\n".join(metadata)))
        
        # Extract embedded images for preview and downloads
        image_count = 0
        if 'extracted_images' not in st.session_state:
            st.session_state.extracted_images = {}
        for rel in doc.part.rels.values():
            if "image" in rel.reltype:
                try:
                    image_part = rel.target_part
                    image_bytes = image_part.blob
                    image_ext = image_part.content_type.split('/')[-1]
                    if image_ext == 'jpeg':
                        image_ext = 'jpg'
                    image_key = f"docx_image_{image_count}"
                    st.session_state.extracted_images[image_key] = {
                        'bytes': image_bytes,
                        'ext': image_ext,
                        'filename': f"image_{image_count}.{image_ext}"
                    }
                    content.append(("EMBEDDED_IMAGE", f"Embedded Image {image_count + 1}: {image_key}"))
                    image_count += 1
                except Exception as e:
                    content.append(("ERROR", f"Could not extract image {image_count + 1}: {e}"))
        if image_count > 0:
            content.append(("METADATA", f"Total Images: {image_count}"))

        # Walk the document body in order and preserve headings, tables, and text.
        current_section_title = None
        current_section_lines = []
        table_count = 0

        def flush_current_section():
            nonlocal current_section_title, current_section_lines
            if current_section_title or current_section_lines:
                if current_section_title:
                    section_text = "\n\n".join(current_section_lines).strip()
                    if section_text:
                        content.append(("TEXT", f"Heading: {current_section_title}\n{section_text}"))
                    else:
                        content.append(("TEXT", f"Heading: {current_section_title}"))
                else:
                    section_text = "\n\n".join(current_section_lines).strip()
                    if section_text:
                        content.append(("TEXT", section_text))
                current_section_title = None
                current_section_lines = []

        def is_docx_heading(para):
            text = para.text.strip()
            if not text:
                return False
            try:
                style_name = (para.style.name or "").lower()
            except Exception:
                style_name = ""
            if "heading" in style_name or style_name.startswith("title") or style_name.startswith("subtitle"):
                return True
            if re.match(r'^\d+(?:\.\d+)*\s+.+', text) and len(text) <= 120:
                return True
            return False

        for element in doc.element.body:
            if element.tag.endswith('}p'):
                paragraph = Paragraph(element, doc)
                paragraph_text = paragraph.text.strip()
                if not paragraph_text:
                    continue
                if is_docx_heading(paragraph):
                    flush_current_section()
                    current_section_title = paragraph_text
                else:
                    current_section_lines.append(paragraph_text)
            elif element.tag.endswith('}tbl'):
                flush_current_section()
                table_count += 1
                table = Table(element, doc)
                table_data = []
                for row in table.rows:
                    row_data = [cell.text.strip() for cell in row.cells]
                    table_data.append(row_data)
                if table_data:
                    table_text = "\n".join([" | ".join(row) for row in table_data])
                    content.append(("TABLE", f"Table {table_count}:\n{table_text}"))

        flush_current_section()

        if table_count > 0:
            content.append(("METADATA", f"Total Tables: {table_count}"))

        if not any(item[0] in ("TEXT", "TABLE", "EMBEDDED_IMAGE", "IMAGE", "METADATA") for item in content):
            paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
            if paragraphs:
                content.append(("TEXT", "\n\n".join(paragraphs)))
    except Exception as e:
        content.append(("ERROR", f"DOCX extraction failed: {str(e)}"))
    return content


def extract_pptx_content(bio):
    """Extract text, tables, and images from PPTX."""
    content = []
    try:
        prs = Presentation(bio)
        
        content.append(("METADATA", f"Total Slides: {len(prs.slides)}"))
        
        # Extract and display images
        image_count = 0
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'image'):
                    try:
                        image_bytes = shape.image.blob
                        image_ext = shape.image.content_type.split('/')[-1]
                        if image_ext == 'jpeg':
                            image_ext = 'jpg'
                        
                        # Create a unique key for the image
                        image_key = f"pptx_image_{image_count}"
                        
                        # Store image data for display
                        if 'extracted_images' not in st.session_state:
                            st.session_state.extracted_images = {}
                        st.session_state.extracted_images[image_key] = {
                            'bytes': image_bytes,
                            'ext': image_ext,
                            'filename': f"slide_image_{image_count}.{image_ext}"
                        }
                        
                        content.append(("EMBEDDED_IMAGE", f"Slide Image {image_count + 1}: {image_key}"))
                        image_count += 1
                    except Exception as e:
                        content.append(("ERROR", f"Could not extract slide image {image_count + 1}: {e}"))
        
        if image_count > 0:
            content.append(("METADATA", f"Total Images: {image_count}"))
        
        # Extract text and tables from each slide
        for i, slide in enumerate(prs.slides):
            slide_content = []
            
            # Extract text shapes
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_content.append(shape.text)
            
            # Extract tables
            for shape in slide.shapes:
                if hasattr(shape, 'table'):
                    table = shape.table
                    table_data = []
                    for row in table.rows:
                        row_data = []
                        for cell in row.cells:
                            cell_text = cell.text.strip()
                            row_data.append(cell_text)
                        table_data.append(row_data)
                    
                    if table_data:
                        table_text = "\n".join([" | ".join(row) for row in table_data])
                        slide_content.append(f"Table:\n{table_text}")
            
            if slide_content:
                content.append(("TEXT", f"Slide {i+1}:\n" + "\n\n".join(slide_content)))
    
    except Exception as e:
        content.append(("ERROR", f"PPTX extraction failed: {str(e)}"))
    
    return content


def extract_xlsx_content(bio):
    """Extract data from all sheets in XLSX."""
    content = []
    try:
        wb = openpyxl.load_workbook(bio, data_only=True)
        
        content.append(("METADATA", f"Workbook contains {len(wb.sheetnames)} sheets: {', '.join(wb.sheetnames)}"))
        
        # Extract data from each sheet
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            sheet_data = []
            
            # Get all rows with data
            for row in sheet.iter_rows(values_only=True):
                if any(cell for cell in row):  # Skip empty rows
                    row_data = [str(cell) if cell is not None else "" for cell in row]
                    sheet_data.append(row_data)
            
            if sheet_data:
                table_text = "\n".join([" | ".join(row) for row in sheet_data])
                content.append(("TABLE", f"Sheet '{sheet_name}':\n{table_text}"))
    
    except Exception as e:
        content.append(("ERROR", f"XLSX extraction failed: {str(e)}"))
    
    return content


def extract_html_content(bio):
    """Extract text and metadata from HTML."""
    content = []
    try:
        html_content = bio.read()
        soup = BeautifulSoup(html_content, "html.parser")
        
        # Extract title
        title = soup.title.string if soup.title else "No title"
        content.append(("METADATA", f"Title: {title}"))
        
        # Extract meta tags
        meta_info = []
        for meta in soup.find_all('meta'):
            if meta.get('name') and meta.get('content'):
                meta_info.append(f"{meta['name']}: {meta['content']}")
        if meta_info:
            content.append(("METADATA", "Meta Tags:\n" + "\n".join(meta_info)))
        
        # Count images
        images = soup.find_all('img')
        if images:
            content.append(("IMAGE", f"{len(images)} images found in HTML"))
        
        # Extract text content
        text = soup.get_text(separator="\n")
        if text.strip():
            content.append(("TEXT", text))
    
    except Exception as e:
        content.append(("ERROR", f"HTML extraction failed: {str(e)}"))
    
    return content


def get_uploaded_file_entry(file_name):
    for file_info in st.session_state.uploaded_files:
        if file_info["name"] == file_name:
            return file_info
    return None


def create_preview_link(file_name, highlight_term=None, page_num=None):
    file_entry = get_uploaded_file_entry(file_name)
    if not file_entry:
        return None
    token = str(uuid.uuid4())
    PREVIEW_TOKENS[token] = {'file_name': file_name, 'timestamp': datetime.now()}
    PREVIEW_STORE[token] = file_entry
    save_preview_data()
    params = [f"preview_token={token}"]
    if highlight_term:
        params.append(f"highlight={urllib.parse.quote_plus(highlight_term)}")
    if page_num is not None:
        params.append(f"page={urllib.parse.quote_plus(str(page_num))}")
    return "./?" + "&".join(params)


def create_heading_anchor(text):
    anchor_text = str(text or "").strip().lower()
    anchor_text = re.sub(r'[^a-z0-9]+', '-', anchor_text)
    anchor_text = re.sub(r'-{2,}', '-', anchor_text).strip('-')
    if not anchor_text:
        anchor_text = 'preview'
    return f"heading-{anchor_text}"


def highlight_for_preview(text, highlight_term=None):
    if not highlight_term:
        return html.escape(text)
    escaped = html.escape(text)
    pattern = re.compile(re.escape(highlight_term), re.IGNORECASE)
    highlighted = pattern.sub(
        lambda m: f"<mark style='background:#fff3a3; padding:0 2px;'>{html.escape(m.group(0))}</mark>",
        escaped
    )
    return highlighted


def render_text_block(text, highlight_term=None, anchor_id=None):
    if not highlight_term:
        return None

    escaped = html.escape(text)
    pattern = re.compile(re.escape(highlight_term), re.IGNORECASE)
    first_match = True

    def replace_match(match):
        nonlocal first_match
        content = html.escape(match.group(0))
        if first_match and anchor_id:
            first_match = False
            return f"<span id='{anchor_id}'></span><mark style='background:#fff3a3; padding:0 2px;'>{content}</mark>"
        first_match = False
        return f"<mark style='background:#fff3a3; padding:0 2px;'>{content}</mark>"

    highlighted = pattern.sub(replace_match, escaped)
    return f"<pre style='white-space: pre-wrap; word-break: break-word; background:#f4f7fb; padding:12px; border-radius:8px; font-family: inherit;'>{highlighted}</pre>"


def render_document_preview(file_name, file_entry=None, highlight_term=None, highlight_page=None):
    st.markdown(f"**Preview: {file_name}**")
    if file_entry is None:
        ensure_file_processed(file_name)
        file_entry = get_uploaded_file_entry(file_name)
    else:
        # Even if file_entry is provided, ensure text is extracted
        if file_name not in st.session_state.file_texts:
            st.session_state.file_texts[file_name] = extract_text(file_name, file_entry["bytes"])
        if file_name.lower().endswith(".xlsx") and file_name not in st.session_state.excel_data_by_file:
            st.session_state.excel_data_by_file[file_name] = extract_excel_data(file_name, file_entry["bytes"])
    
    if not file_entry:
        st.warning("File preview is unavailable.")
        return

    file_name_lower = file_name.lower()
    image_download_items = []
    table_download_items = []

    # Special handling for PDF files: render actual page images for a true preview
    if file_name_lower.endswith(".pdf"):
        with st.spinner("Rendering PDF preview..."):
            try:
                pdf_bio = BytesIO(file_entry["bytes"])
                with pdfplumber.open(pdf_bio) as pdf:
                    st.markdown(f"**PDF Pages: {len(pdf.pages)}**")
                    highlight_found = False
                    selected_page_index = None
                    if highlight_page is not None and 1 <= highlight_page <= len(pdf.pages):
                        selected_page_index = highlight_page - 1
                    for i, page in enumerate(pdf.pages):
                        if selected_page_index is not None and i != selected_page_index:
                            continue
                        try:
                            page_text = page.extract_text() or ""
                            page_anchor_id = None
                            if highlight_term and highlight_term.lower() in page_text.lower():
                                page_anchor_id = create_heading_anchor(highlight_term)
                                highlight_found = True
                            if page_anchor_id:
                                st.markdown(f"<div id='{page_anchor_id}'></div>", unsafe_allow_html=True)

                            if selected_page_index is not None:
                                st.markdown(f"**Showing page {highlight_page} only.**")
                            page_image = page.to_image(resolution=150)
                            image_bytes_io = BytesIO()
                            page_image.original.save(image_bytes_io, format="PNG")
                            image_bytes = image_bytes_io.getvalue()

                            st.image(image_bytes, caption=f"Page {i+1}", use_container_width=True)
                            if page_anchor_id and highlight_term:
                                st.markdown("### Highlighted page text", unsafe_allow_html=True)
                                st.markdown(render_text_block(page_text, highlight_term, anchor_id=None), unsafe_allow_html=True)
                            image_download_items.append({
                                "label": f"Download Page {i+1} as PNG",
                                "data": image_bytes,
                                "file_name": f"{os.path.splitext(file_name)[0]}_page_{i+1}.png",
                                "mime": "image/png",
                                "key": f"download_pdf_page_{file_name}_{i}"
                            })

                            tables = page.extract_tables()
                            if tables:
                                for j, table in enumerate(tables):
                                    if table and any(any(cell for cell in row) for row in table):
                                        table_png = table_to_png_bytes(table, title=f"Page {i+1} Table {j+1}")
                                        st.image(table_png, caption=f"Page {i+1} Table {j+1}", use_container_width=True)
                                        table_download_items.append({
                                            "label": f"📥 Download Table {j+1} as PNG",
                                            "data": table_png,
                                            "file_name": f"{os.path.splitext(file_name)[0]}_page_{i+1}_table_{j+1}.png",
                                            "mime": "image/png",
                                            "key": f"download_pdf_table_{file_name}_{i}_{j}"
                                        })
                        except Exception as page_err:
                            st.warning(f"Could not render page {i+1} as image: {page_err}")
                            page_text = page.extract_text() or ""
                            if page_text.strip():
                                page_anchor_id = None
                                if highlight_term and highlight_term.lower() in page_text.lower():
                                    page_anchor_id = create_heading_anchor(highlight_term)
                                if page_anchor_id:
                                    st.markdown(f"<div id='{page_anchor_id}'></div>", unsafe_allow_html=True)
                                st.markdown(f"#### Page {i+1} Text")
                                if highlight_term:
                                    st.markdown(render_text_block(page_text, highlight_term, anchor_id=page_anchor_id), unsafe_allow_html=True)
                                else:
                                    st.code(page_text, language="text")
                    if highlight_term and not highlight_found and selected_page_index is None:
                        if file_name not in st.session_state.file_texts:
                            st.session_state.file_texts[file_name] = extract_text(file_name, file_entry["bytes"])
                        full_content = st.session_state.file_texts.get(file_name, "")
                        anchor_id = create_heading_anchor(highlight_term)
                        st.markdown(f"<div id='{anchor_id}'></div>", unsafe_allow_html=True)
                        st.markdown("### Highlighted Text")
                        st.markdown(render_text_block(full_content, highlight_term, anchor_id=None), unsafe_allow_html=True)

                    if image_download_items:
                        with st.expander("🖼️ Image Downloads", expanded=False):
                            for item in image_download_items:
                                st.download_button(
                                    label=item["label"],
                                    data=item["data"],
                                    file_name=item["file_name"],
                                    mime=item["mime"],
                                    key=item["key"]
                                )
                    if table_download_items:
                        with st.expander("📊 Table Downloads", expanded=False):
                            for item in table_download_items:
                                st.download_button(
                                    label=item["label"],
                                    data=item["data"],
                                    file_name=item["file_name"],
                                    mime=item["mime"],
                                    key=item["key"]
                                )
                    if highlight_term and not highlight_found and selected_page_index is None:
                        if file_name not in st.session_state.file_texts:
                            st.session_state.file_texts[file_name] = extract_text(file_name, file_entry["bytes"])
                        full_content = st.session_state.file_texts.get(file_name, "")
                        if full_content.strip():
                            anchor_id = create_heading_anchor(highlight_term)
                            st.markdown(f"<div id='{anchor_id}'></div>", unsafe_allow_html=True)
                            st.markdown("### Highlighted Text")
                            st.markdown(render_text_block(full_content, highlight_term, anchor_id=anchor_id), unsafe_allow_html=True)
                    return
            except Exception as e:
                st.error(f"Could not render PDF preview: {e}")
                st.info("Falling back to text-based document preview.")

    # Special handling for images
    if file_name_lower.endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp")):
        with st.spinner("Loading preview..."):
            mime_type = "application/octet-stream"
            try:
                st.image(file_entry["bytes"], caption=file_name, use_container_width=True)
                png_bytes = image_bytes_to_png_bytes(file_entry["bytes"])
                # Determine MIME type
                ext = file_name_lower.split('.')[-1]
                mime_type = f"image/{ext}"
                if ext == "jpg":
                    mime_type = "image/jpeg"
                elif ext == "svg":
                    mime_type = "image/svg+xml"

                st.download_button(
                    label="Download Image",
                    data=file_entry["bytes"],
                    file_name=file_name,
                    mime=mime_type,
                    key=f"download_image_{file_name}"
                )
                st.download_button(
                    label="Download as PNG",
                    data=png_bytes,
                    file_name=f"{os.path.splitext(file_name)[0]}.png",
                    mime="image/png",
                    key=f"download_image_png_{file_name}"
                )
            except Exception as e:
                st.error(f"Could not display image: {e}")
                st.download_button(
                    label="Download Image File",
                    data=file_entry["bytes"],
                    file_name=file_name,
                    mime=mime_type,
                    key=f"download_image_fallback_{file_name}"
                )
        return

    # Special handling for Excel files (show as table)
    if file_name_lower.endswith(".xlsx"):
        with st.spinner("Loading preview..."):
            data = st.session_state.excel_data_by_file.get(file_name, [])
            if data:
                st.markdown("### Excel Preview")
                preview_df = pd.DataFrame(data).head(20)
                st.dataframe(preview_df, use_container_width=True, hide_index=True)
                st.download_button(
                    label="Download table as CSV",
                    data=preview_df.to_csv(index=False),
                    file_name=f"{os.path.splitext(file_name)[0]}_preview.csv",
                    mime="text/csv",
                    key=f"download_excel_csv_{file_name}"
                )
                try:
                    table_png = table_to_png_bytes(
                        dataframe_to_table_rows(preview_df),
                        title=f"{file_name} Preview"
                    )
                    st.download_button(
                        label="Download table as PNG",
                        data=table_png,
                        file_name=f"{os.path.splitext(file_name)[0]}_preview.png",
                        mime="image/png",
                        key=f"download_excel_png_{file_name}"
                    )
                except Exception:
                    pass
            else:
                st.info("No preview data available for this spreadsheet.")
        return

    # For all other files, show comprehensive extracted content
    with st.spinner("Loading preview..."):
        # Ensure text extraction has run
        if file_name not in st.session_state.file_texts:
            st.session_state.file_texts[file_name] = extract_text(file_name, file_entry["bytes"])
        
        full_content = st.session_state.file_texts.get(file_name, "")
        
        if not full_content.strip():
            st.warning("No content could be extracted from this file. This might indicate an issue with the file format or content extraction.")
            # Try to show basic file info
            file_size = len(file_entry["bytes"])
            st.info(f"File size: {file_size} bytes")
            st.info(f"File type: {file_name.split('.')[-1].upper()}")
            return

        # Parse the content into sections
        sections = parse_extracted_content(full_content)
        
        # For PDF files, if parsing doesn't produce good content, show raw text
        if file_name_lower.endswith('.pdf') and sections:
            # Check if we have meaningful text content
            has_meaningful_text = any(
                section_type == "TEXT" and section_content.strip()
                for section_type, _, section_content in sections
            )
            if not has_meaningful_text:
                # Fall back to showing raw content
                st.markdown("### Document Content")
                anchor_id = create_heading_anchor(highlight_term) if highlight_term else None
                with st.expander("Show extracted text", expanded=True):
                    if highlight_term:
                        st.markdown(render_text_block(full_content.strip(), highlight_term, anchor_id=anchor_id), unsafe_allow_html=True)
                    else:
                        st.code(full_content.strip(), language="text")
                return
        
        if not sections:
            st.warning("Content was extracted but could not be parsed into displayable sections.")
            anchor_id = create_heading_anchor(highlight_term) if highlight_term else None
            if highlight_term:
                st.markdown(render_text_block(full_content[:1000] + ("..." if len(full_content) > 1000 else ""), highlight_term, anchor_id=anchor_id), unsafe_allow_html=True)
            else:
                st.code(full_content[:1000] + ("..." if len(full_content) > 1000 else ""), language="text")
            return
        
        # Display each section
        for section_type, section_title, section_content in sections:
            if section_type == "METADATA":
                with st.expander(f"📋 {section_title}", expanded=False):
                    if section_content.strip():
                        if highlight_term:
                            st.markdown(render_text_block(section_content, highlight_term), unsafe_allow_html=True)
                        else:
                            st.code(section_content, language="text")
                    else:
                        st.info("No metadata available")
            elif section_type == "TEXT":
                section_anchor_id = create_heading_anchor(section_title)
                st.markdown(f"<div id='{section_anchor_id}'></div>", unsafe_allow_html=True)
                st.markdown(f"### {section_title}")
                if section_content.strip():
                    if len(section_content) > 2000:
                        with st.expander("Show text content", expanded=False):
                            if highlight_term:
                                st.markdown(render_text_block(section_content, highlight_term, anchor_id=None), unsafe_allow_html=True)
                            else:
                                st.code(section_content, language="text")
                    else:
                        if highlight_term:
                            st.markdown(render_text_block(section_content, highlight_term, anchor_id=None), unsafe_allow_html=True)
                        else:
                            st.code(section_content, language="text")
                else:
                    st.info("No text content available for this section.")
            elif section_type == "TABLE":
                with st.expander(f"📊 {section_title}", expanded=True):
                    # Try to parse table and display as dataframe
                    try:
                        lines = section_content.strip().split('\n')
                        if lines:
                            # Parse table data
                            table_data = []
                            for line in lines:
                                if ' | ' in line:
                                    row = [cell.strip() for cell in line.split(' | ')]
                                    table_data.append(row)
                            
                            if table_data:
                                df = pd.DataFrame(table_data[1:] if len(table_data) > 1 else table_data, 
                                                columns=table_data[0] if len(table_data) > 1 else None)
                                st.dataframe(df, use_container_width=True, hide_index=True)
                                
                                # Add download button for table as CSV
                                csv_data = df.to_csv(index=False)
                                st.download_button(
                                    label="Download as CSV",
                                    data=csv_data,
                                    file_name=f"{section_title.replace(' ', '_')}.csv",
                                    mime="text/csv",
                                    key=f"download_table_{file_name}_{section_title}"
                                )
                                try:
                                    table_png = table_to_png_bytes(table_data, title=section_title)
                                    table_download_items.append({
                                        "label": f"Download {section_title} as PNG",
                                        "data": table_png,
                                        "file_name": f"{section_title.replace(' ', '_')}.png",
                                        "mime": "image/png",
                                        "key": f"download_table_png_{file_name}_{section_title}"
                                    })
                                except Exception:
                                    pass
                            else:
                                if highlight_term:
                                    st.markdown(render_text_block(section_content, highlight_term), unsafe_allow_html=True)
                                else:
                                    st.code(section_content, language="text")
                    except Exception:
                        if highlight_term:
                            st.markdown(render_text_block(section_content, highlight_term), unsafe_allow_html=True)
                        else:
                            st.code(section_content, language="text")
            elif section_type == "EMBEDDED_IMAGE":
                # Display embedded image
                image_key = section_content.split(": ")[-1] if ": " in section_content else section_content
                if 'extracted_images' in st.session_state and image_key in st.session_state.extracted_images:
                    image_data = st.session_state.extracted_images[image_key]
                    try:
                        st.image(image_data['bytes'], caption=image_data['filename'], use_container_width=True)
                        # Add download button
                        mime_type = f"image/{image_data['ext']}"
                        if image_data['ext'] == "jpg":
                            mime_type = "image/jpeg"
                        st.download_button(
                            label="Download Image",
                            data=image_data['bytes'],
                            file_name=image_data['filename'],
                            mime=mime_type,
                            key=f"download_embedded_{image_key}"
                        )
                        try:
                            png_bytes = image_bytes_to_png_bytes(image_data['bytes'])
                            st.download_button(
                                label="Download as PNG",
                                data=png_bytes,
                                file_name=f"{os.path.splitext(image_data['filename'])[0]}.png",
                                mime="image/png",
                                key=f"download_embedded_png_{image_key}"
                            )
                        except Exception:
                            pass
                    except Exception as e:
                        st.error(f"Could not display image: {e}")
                else:
                    st.info(f"🖼️ {section_content}")
            elif section_type == "IMAGE":
                st.info(f"🖼️ {section_content}")
            elif section_type == "ERROR":
                st.error(f"❌ {section_content}")
            elif section_type == "UNSUPPORTED":
                st.warning(f"⚠️ {section_content}")

        if image_download_items:
            with st.expander("🖼️ Image Downloads", expanded=False):
                for item in image_download_items:
                    st.download_button(
                        label=item["label"],
                        data=item["data"],
                        file_name=item["file_name"],
                        mime=item["mime"],
                        key=item["key"]
                    )
        if table_download_items:
            with st.expander("📊 Table Downloads", expanded=False):
                for item in table_download_items:
                    st.download_button(
                        label=item["label"],
                        data=item["data"],
                        file_name=item["file_name"],
                        mime=item["mime"],
                        key=item["key"]
                    )


def parse_extracted_content(content):
    """Parse the extracted content into sections for display."""
    sections = []
    lines = content.split('\n')
    current_section = None
    current_content = []

    def flush_section():
        nonlocal current_section, current_content
        if not current_section:
            return

        section_type, section_title, section_value = current_section
        if section_type in ("TEXT", "TABLE"):
            sections.append((section_type, section_title, '\n'.join(current_content).strip()))
        else:
            final_value = section_value
            if current_content:
                final_value = '\n'.join(current_content).strip()
            sections.append((section_type, section_title, final_value))

        current_section = None
        current_content = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Check for section markers
        if line.startswith('TABLE:'):
            # Save previous section
            flush_section()
            
            # Start new table section
            current_section = ("TABLE", "Table Content", "")
            current_content = []
            
        elif line.startswith('[IMAGE:'):
            # Save previous section
            flush_section()
            
            # Add image info
            sections.append(("IMAGE", "Images", line))
            current_section = None
            current_content = []
            
        elif line.startswith('[EMBEDDED_IMAGE:'):
            # Save previous section
            flush_section()
            
            # Add embedded image info
            image_info = line.replace('[EMBEDDED_IMAGE: ', '').replace(']', '')
            sections.append(("EMBEDDED_IMAGE", "Embedded Image", image_info))
            current_section = None
            current_content = []
            
        elif line.startswith('PDF Metadata:') or line.startswith('Document Metadata:') or line.startswith('Meta Tags:') or line.startswith('Title:') or 'Pages:' in line or 'Slides:' in line or 'sheets:' in line:
            # Save previous section
            flush_section()
            
            # Start metadata section
            if 'Metadata:' in line or 'Tags:' in line:
                section_title = line.split(':')[0] + " Information"
            else:
                section_title = "Document Information"
            
            current_section = ("METADATA", section_title, line)
            current_content = [line]
            
        elif line.startswith('Heading:'):
            # Save previous section
            flush_section()
            
            heading_title = line.replace('Heading:', '', 1).strip()
            current_section = ("TEXT", heading_title or "Heading", "")
            current_content = []

        elif line.startswith('Page ') and 'Text:' in line:
            # Save previous section
            flush_section()
            
            # Start new text section
            current_section = ("TEXT", f"Page {line.split()[1]} Content", "")
            current_content = []
            
        elif line.startswith('Slide ') and ':' in line:
            # Save previous section
            flush_section()
            
            # Start new slide section
            slide_num = line.split(':')[0]
            current_section = ("TEXT", f"{slide_num} Content", "")
            current_content = []
            
        elif line.startswith('Sheet ') and ':' in line:
            # Save previous section
            flush_section()
            
            # Start new sheet section
            sheet_name = line.split(':')[0].replace("'", "")
            current_section = ("TABLE", f"{sheet_name} Data", "")
            current_content = []
            
        elif current_section:
            # Add to current section
            if current_section[0] == "METADATA":
                current_content.append(line)
                current_section = (current_section[0], current_section[1], '\n'.join(current_content))
            else:
                current_content.append(line)
        else:
            # Start default text section
            if not current_section:
                current_section = ("TEXT", "Document Content", "")
                current_content = []
            current_content.append(line)
    
    # Save final section
    flush_section()
    
    # If no sections were created but we have content, create a default text section
    if not sections and content.strip():
        sections.append(("TEXT", "Document Content", content.strip()))
    
    return sections


@st.cache_data(show_spinner=False)
def build_summary_download_assets(file_name, file_bytes):
    file_name_lower = file_name.lower()
    image_items = []
    table_items = []

    try:
        if file_name_lower.endswith(".pdf"):
            with pdfplumber.open(BytesIO(file_bytes)) as pdf:
                for page_index, page in enumerate(pdf.pages, start=1):
                    for image_index, image_info in enumerate(page.images or [], start=1):
                        try:
                            bbox = (
                                image_info.get("x0", 0),
                                image_info.get("top", 0),
                                image_info.get("x1", 0),
                                image_info.get("bottom", 0)
                            )
                            if bbox[0] < bbox[2] and bbox[1] < bbox[3]:
                                image_items.append({
                                    "label": f"{file_name} - Page {page_index} Image {image_index}",
                                    "data": crop_pdf_region_to_png(page, bbox),
                                    "file_name": f"{os.path.splitext(file_name)[0]}_page_{page_index}_image_{image_index}.png",
                                    "mime": "image/png"
                                })
                        except Exception:
                            pass

                    for table_index, table in enumerate(page.extract_tables() or [], start=1):
                        if table and any(any(cell for cell in row) for row in table):
                            try:
                                table_items.append({
                                    "label": f"{file_name} - Page {page_index} Table {table_index}",
                                    "data": table_to_png_bytes(table, title=f"Page {page_index} Table {table_index}"),
                                    "file_name": f"{os.path.splitext(file_name)[0]}_page_{page_index}_table_{table_index}.png",
                                    "mime": "image/png"
                                })
                            except Exception:
                                pass

        elif file_name_lower.endswith(".docx"):
            doc = docx.Document(BytesIO(file_bytes))
            image_index = 1
            for rel in doc.part.rels.values():
                if "image" in rel.reltype:
                    try:
                        image_bytes = rel.target_part.blob
                        image_items.append({
                            "label": f"{file_name} - Image {image_index}",
                            "data": image_bytes_to_png_bytes(image_bytes),
                            "file_name": f"{os.path.splitext(file_name)[0]}_image_{image_index}.png",
                            "mime": "image/png"
                        })
                        image_index += 1
                    except Exception:
                        pass

            for table_index, table in enumerate(doc.tables, start=1):
                table_data = [[cell.text.strip() for cell in row.cells] for row in table.rows]
                if table_data:
                    try:
                        table_items.append({
                            "label": f"{file_name} - Table {table_index}",
                            "data": table_to_png_bytes(table_data, title=f"Table {table_index}"),
                            "file_name": f"{os.path.splitext(file_name)[0]}_table_{table_index}.png",
                            "mime": "image/png"
                        })
                    except Exception:
                        pass

        elif file_name_lower.endswith(".pptx"):
            prs = Presentation(BytesIO(file_bytes))
            image_index = 1
            table_index = 1
            for slide_index, slide in enumerate(prs.slides, start=1):
                for shape in slide.shapes:
                    if hasattr(shape, "image"):
                        try:
                            image_items.append({
                                "label": f"{file_name} - Slide {slide_index} Image {image_index}",
                                "data": image_bytes_to_png_bytes(shape.image.blob),
                                "file_name": f"{os.path.splitext(file_name)[0]}_slide_{slide_index}_image_{image_index}.png",
                                "mime": "image/png"
                            })
                            image_index += 1
                        except Exception:
                            pass
                    if hasattr(shape, "table"):
                        try:
                            table_data = [
                                [cell.text.strip() for cell in row.cells]
                                for row in shape.table.rows
                            ]
                            table_items.append({
                                "label": f"{file_name} - Slide {slide_index} Table {table_index}",
                                "data": table_to_png_bytes(table_data, title=f"Slide {slide_index} Table {table_index}"),
                                "file_name": f"{os.path.splitext(file_name)[0]}_slide_{slide_index}_table_{table_index}.png",
                                "mime": "image/png"
                            })
                            table_index += 1
                        except Exception:
                            pass

        elif file_name_lower.endswith(".xlsx"):
            data = extract_excel_data(file_name, file_bytes)
            if data:
                preview_df = pd.DataFrame(data).head(20)
                table_items.append({
                    "label": f"{file_name} - Preview Table",
                    "data": table_to_png_bytes(dataframe_to_table_rows(preview_df), title=f"{file_name} Preview"),
                    "file_name": f"{os.path.splitext(file_name)[0]}_preview.png",
                    "mime": "image/png"
                })

        elif file_name_lower.endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp")):
            image_items.append({
                "label": f"{file_name} - Image",
                "data": image_bytes_to_png_bytes(file_bytes),
                "file_name": f"{os.path.splitext(file_name)[0]}.png",
                "mime": "image/png"
            })
    except Exception:
        return {"images": [], "tables": []}

    return {"images": image_items, "tables": table_items}


def render_extracted_assets_preview(file_name, file_entry):
    st.markdown(f"**Preview: {file_name}**")
    assets = build_summary_download_assets(file_name, file_entry["bytes"])
    image_items = assets.get("images", [])
    table_items = assets.get("tables", [])

    if not image_items and not table_items:
        st.info("No extractable images or tables were found in this file.")
        return

    if image_items:
        st.markdown("### Extracted Images")
        for index, item in enumerate(image_items):
            st.image(item["data"], caption=item["label"], use_container_width=True)
            st.download_button(
                label=f"Download {item['label']} as PNG",
                data=item["data"],
                file_name=item["file_name"],
                mime=item["mime"],
                key=f"preview_image_download_{index}_{item['file_name']}"
            )

    if table_items:
        st.markdown("### Extracted Tables")
        for index, item in enumerate(table_items):
            st.image(item["data"], caption=item["label"], use_container_width=True)
            st.download_button(
                label=f"Download {item['label']} as PNG",
                data=item["data"],
                file_name=item["file_name"],
                mime=item["mime"],
                key=f"preview_table_download_{index}_{item['file_name']}"
            )


# Preview route handling:
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
    token_data = PREVIEW_TOKENS.get(preview_token)
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
        st.title("Preview Error")
        st.error(f"Preview token '{preview_token}' not found or expired.")
        st.write("Debug info:")
        st.write(f"Available tokens: {list(PREVIEW_TOKENS.keys())}")
        st.write(f"Query params: {query_params}")
        st.stop()

if preview_file_from_url:
    preview_entry = PREVIEW_STORE.get(preview_token)
    st.title("📄 Document Preview")
    if preview_entry is not None:
        st.markdown(f"### {preview_entry['name']}")
        st.markdown("---")
        render_document_preview(preview_entry['name'], file_entry=preview_entry, highlight_term=highlight_term, highlight_page=preview_page)
        show_assets = st.checkbox("Show extracted asset previews", value=False)
        if show_assets:
            st.markdown("---")
            st.markdown("### Extracted Assets")
            render_extracted_assets_preview(preview_entry['name'], file_entry=preview_entry)
    else:
        st.error("Preview file not found in the preview store. Please return to the app and click preview again.")
    st.stop()


# -------------------------------
# FILE UPLOAD & MANAGEMENT (SIDEBAR)
# -------------------------------
# Sidebar area:
# This block manages login state, file upload/delete, global file selection, and
# preview launch links. Files selected here become available to the individual tabs.
with st.sidebar:
    if st.session_state.is_authenticated:
        creator_timestamp = None
        if st.session_state.user_role == "creator" and st.session_state.login_history:
            creator_entries = [
                entry for entry in st.session_state.login_history
                if entry.get("username") == st.session_state.logged_in_username and entry.get("role") == "creator"
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
                    background: #f8fbff;
                    border: 1px solid #d7e3f4;
                    color: #173152;
                }
                [data-testid="stSidebar"] [class*="st-key-select_file_"] button[kind="primary"] {
                    background: #eaf2ff;
                    border: 1px solid #1f4f91;
                    color: #173152;
                }
                .file-icon-button {
                    display: inline-flex;
                    justify-content: center;
                    align-items: center;
                    width: 38px;
                    height: 38px;
                    border-radius: 12px;
                    background: transparent;
                    color: #1f4f91;
                    text-decoration: none;
                    font-size: 18px;
                }
                .file-icon-button:hover {
                    background: #eaf2ff;
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

        st.header("Upload Documents")
        st.info("1) Upload files." \
        " 2) Click the file cards you need. " \
        "3) Switch tabs and work with selected files.")
        new_files = st.file_uploader(
            "Upload PDF, DOCX, TXT, PPTX, XLSX, HTML, CAPL, Images",
            type=["pdf", "docx", "txt", "pptx", "xlsx", "html", "htm", "capl", "can", "png", "jpg", "jpeg", "gif", "bmp", "webp"],
            accept_multiple_files=True,
            key=f"file_uploader_{st.session_state.file_uploader_key}"
        )

        if new_files:
            with st.spinner("Loading uploaded files..."):
                existing_names = {f["name"] for f in st.session_state.uploaded_files}
                uploaded_new_file = False
                for file in new_files:
                    if file.name not in existing_names:
                        file_bytes = file.read()
                        st.session_state.uploaded_files.append({"name": file.name, "bytes": file_bytes})
                        uploaded_new_file = True
                if uploaded_new_file:
                    st.session_state.messages = []
                    st.session_state.chat_summary_downloads = {"images": [], "tables": []}
                    st.session_state.chat_file_selection = []
                    st.success("✅ New files uploaded. Chat history has been cleared.")

        st.markdown("---")
        st.markdown("### Uploaded files")
        
        for idx, file_dict in enumerate(st.session_state.uploaded_files[:]):
            cols = st.columns([0.65, 0.18, 0.17], vertical_alignment="center")
            with cols[0]:
                file_name = file_dict["name"]
                is_selected = file_name in st.session_state.selected_files
                button_label = file_name if not is_selected else f"Selected: {file_name}"
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
                token = str(uuid.uuid4())
                PREVIEW_TOKENS[token] = {'file_name': file_name, 'timestamp': datetime.now()}
                PREVIEW_STORE[token] = file_dict
                save_preview_data()
                st.markdown(
                    f"<a href='./?preview_token={token}' target='_blank' class='file-icon-button' title='Preview {html.escape(file_name)}'>👁️</a>",
                    unsafe_allow_html=True,
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

                    if st.session_state.capl_last_analyzed_file == deleted_name:
                        st.session_state.capl_last_analyzed_file = None
                        st.session_state.capl_last_issues = None
                    st.rerun()
        st.markdown("*Selected files above are available across all tabs.*")
        st.markdown("---")
        if st.button("Clear All Files"):
            for key in ["uploaded_files", "selected_files", "file_texts", "excel_data_by_file", "vector_stores",
                        "messages"]:
                st.session_state[key].clear()
            st.session_state.chat_summary_downloads = {"images": [], "tables": []}
            st.session_state.chat_file_selection = []
            st.session_state.capl_last_analyzed_file = None
            st.session_state.capl_last_issues = None
            st.session_state.file_uploader_key += 1
            st.rerun()

        if st.session_state.user_role == "creator":
            st.markdown("---")
            st.subheader("Creator Login History")
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
@st.cache_data(show_spinner=False)
def extract_excel_data(file_name, file_bytes):
    data = []
    bio = BytesIO(file_bytes)
    file_name_lower = file_name.lower()
    try:
        if file_name_lower.endswith(".xlsx"):
            wb = openpyxl.load_workbook(bio, data_only=True)
            for sheet in wb:
                headers = None
                for i, row in enumerate(sheet.iter_rows(values_only=True)):
                    if i == 0:
                        headers = list(row)
                    else:
                        if row and any(cell is not None for cell in row):
                            data.append(dict(zip(headers, row)))
    except Exception:
        data = []
    return data


# -------------------------------
# PROCESS FILES & BUILD VECTOR STORES
# -------------------------------
# AI/vector helpers:
# These are mainly used by the Chat tab for semantic retrieval and LLM answers.
# They also centralize file preprocessing so each tab can rely on the same cache.
@st.cache_data(show_spinner=False)
def create_vector_store(text):
    splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=100)
    chunks = splitter.split_text(text[:50000])  # limit size for speed
    emb = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")
    return FAISS.from_texts(chunks, emb)


@st.cache_resource(show_spinner=False)
def load_llm():
    try:
        from transformers import pipeline
    except Exception:
        st.warning(
            "LLM is unavailable because transformers could not be imported (torchvision unmet).\nInstall torch & torchvision if you want AI features.")
        return None

    candidate_tasks = ["text2text-generation", "text-generation", "image-text-to-text", "table-question-answering"]
    task_errors = []

    for task in candidate_tasks:
        try:
            pipe = pipeline(task, model="google/flan-t5-small", max_new_tokens=128, return_full_text=False)
            st.session_state.llm_task = task
            return HuggingFacePipeline(pipeline=pipe)
        except Exception as e:
            task_errors.append((task, str(e)))

    # Only show a single warning if no task could be initialized
    if task_errors:
        st.warning("LLM initialization failed for all candidate tasks; AI features are unavailable.")
        st.session_state.llm_task = None

    return None


def ensure_files_processed(file_names):
    for file_name in file_names:
        ensure_file_processed(file_name)


def process_selected_files():
    """Fully extract all currently selected files so every tab uses the same data."""
    ensure_files_processed(st.session_state.selected_files)


def get_selection_signature(file_names):
    digest = hashlib.md5()
    for file_name in sorted(file_names):
        digest.update(file_name.encode("utf-8"))
        digest.update(st.session_state.file_texts.get(file_name, "").encode("utf-8"))
    return f"combined::{digest.hexdigest()}"


SUMMARY_STOPWORDS = {
    "the", "and", "for", "with", "that", "this", "from", "are", "was", "were", "into", "your", "have",
    "has", "had", "not", "but", "you", "all", "can", "will", "use", "using", "used", "how", "what",
    "when", "where", "which", "while", "into", "more", "than", "their", "there", "about", "after",
    "before", "within", "without", "each", "page", "pages", "table", "tables", "image", "images",
    "document", "content", "metadata", "information", "product", "file", "text"
}


@st.cache_data(show_spinner=False)
def get_document_asset_counts(file_name, file_bytes, extracted_text):
    file_name_lower = file_name.lower()
    page_count = 0
    table_count = 0
    image_count = 0

    if file_name_lower.endswith(".pdf"):
        try:
            with pdfplumber.open(BytesIO(file_bytes)) as pdf:
                page_count = len(pdf.pages)
                for page in pdf.pages:
                    tables = page.extract_tables() or []
                    table_count += sum(1 for table in tables if table and any(any(cell for cell in row) for row in table))
                    image_count += len(getattr(page, "images", []) or [])
        except Exception:
            page_match = re.search(r"Total Pages:\s*(\d+)", extracted_text)
            page_count = int(page_match.group(1)) if page_match else 0
            table_count = len(re.findall(r"Page \d+ Table \d+:", extracted_text))
            image_count = len(re.findall(r"\[IMAGE:", extracted_text))
    elif file_name_lower.endswith(".pptx"):
        slide_match = re.search(r"Total Slides:\s*(\d+)", extracted_text)
        page_count = int(slide_match.group(1)) if slide_match else 0
        table_count = len(re.findall(r"\bTable:\n", extracted_text))
        image_count = len(re.findall(r"\[EMBEDDED_IMAGE:", extracted_text))
    elif file_name_lower.endswith(".docx"):
        table_count = len(re.findall(r"Table \d+:", extracted_text))
        image_count = len(re.findall(r"\[EMBEDDED_IMAGE:", extracted_text))
    elif file_name_lower.endswith(".xlsx"):
        sheet_match = re.search(r"Workbook contains (\d+) sheets", extracted_text)
        page_count = int(sheet_match.group(1)) if sheet_match else 0
        table_count = len(re.findall(r"Sheet '.*?':", extracted_text))
    elif file_name_lower.endswith((".html", ".htm")):
        image_match = re.search(r"(\d+) images found in HTML", extracted_text)
        image_count = int(image_match.group(1)) if image_match else 0

    return page_count, image_count, table_count


# Chat summary helper:
# Used only in the Chat tab after summarize/analyze actions to expose extracted
# images and tables as downloads for the files currently selected in chat.
def render_chat_summary_downloads():
    downloads = st.session_state.get("chat_summary_downloads", {"images": [], "tables": []})
    image_items = downloads.get("images", [])
    table_items = downloads.get("tables", [])

    if not image_items and not table_items:
        return

    st.markdown("### Summary Downloads")

    if image_items:
        with st.expander("🖼️ Image PNG Downloads", expanded=False):
            for index, item in enumerate(image_items):
                st.download_button(
                    label=item["label"],
                    data=item["data"],
                    file_name=item["file_name"],
                    mime=item["mime"],
                    key=f"chat_summary_image_{index}_{item['file_name']}"
                )

    if table_items:
        with st.expander("📊 Table PNG Downloads", expanded=False):
            for index, item in enumerate(table_items):
                st.download_button(
                    label=item["label"],
                    data=item["data"],
                    file_name=item["file_name"],
                    mime=item["mime"],
                    key=f"chat_summary_table_{index}_{item['file_name']}"
                )


def build_detailed_document_summary(file_name, file_bytes, text):
    lines = [line.strip() for line in str(text).splitlines() if line.strip()]
    words = re.findall(r"\w+", str(text))
    title_match = re.search(r"Title:\s*(.+)", str(text))
    title = title_match.group(1).strip() if title_match and title_match.group(1).strip() else file_name

    keyword_counts = Counter(
        word.lower()
        for word in words
        if len(word) > 3 and word.lower() not in SUMMARY_STOPWORDS and not word.isdigit()
    )
    keywords = ", ".join(word.title() for word, _ in keyword_counts.most_common(6)) or "Not available"

    page_count, image_count, table_count = get_document_asset_counts(file_name, file_bytes, str(text))

    ignored_prefixes = (
        "pdf metadata:", "document metadata:", "meta tags:", "total pages:", "total slides:",
        "workbook contains", "error:", "[image:", "[embedded_image:"
    )
    key_lines = []
    seen_lines = set()
    for line in lines:
        if line.lower().startswith(ignored_prefixes):
            continue
        if len(line) < 4:
            continue
        if line in seen_lines:
            continue
        seen_lines.add(line)
        key_lines.append(line)
        if len(key_lines) == 5:
            break

    preview_text = " ".join(key_lines[:3] if key_lines else lines[:3])[:500]

    escaped_file_name = html.escape(file_name)
    escaped_title = html.escape(title)
    escaped_keywords = html.escape(keywords)
    key_lines_html = "".join(
        f"<div>{html.escape(line)}</div>"
        for line in key_lines
    ) if key_lines else "<div>No key content lines could be extracted.</div>"
    preview_html = html.escape(f"{preview_text}..." if preview_text else "No readable preview available.")

    return f"""
    <div style="margin-bottom:16px;">
        <div style="font-weight:600; color:#173152; margin-bottom:8px;">📄 {escaped_file_name}</div>
        <div style="font-weight:600; margin:10px 0 6px 0;">Key Information:</div>
        <div>Title: {escaped_title}</div>
        <div>Keywords: {escaped_keywords}</div>
        {key_lines_html}
        <div style="font-weight:600; margin:12px 0 6px 0;">Document Statistics:</div>
        <div>Total characters: {len(str(text))}</div>
        <div>Estimated words: {len(words)}</div>
        <div>Content lines: {len(lines)}</div>
        <div>Pages/Sections: {page_count}</div>
        <div>Images found: {image_count}</div>
        <div>Tables found: {table_count}</div>
        <div>Downloadable preview assets: Images {image_count}, Tables {table_count}</div>
        <div style="font-weight:600; margin:12px 0 6px 0;">Content Preview:</div>
        <div>{preview_html}</div>
    </div>
    """


# Document structure helpers:
# Used by the Chat "overview" flow and preview links to identify headings,
# table-of-contents entries, and likely page numbers inside uploaded documents.
def extract_page_text(text, page_number=1):
    text = str(text)
    pattern = rf"Page {page_number}\s+Text:\s*(.*?)(?=Page \d+\s+Text:|\Z)"
    match = re.search(pattern, text, re.S | re.IGNORECASE)
    if match:
        return match.group(1).strip()

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return "\n".join(lines[:80])


def find_heading_page_number(text, heading):
    text = str(text)
    lines = [line for line in text.splitlines()]
    heading_pattern = re.escape(str(heading).strip())
    for index, line in enumerate(lines):
        if re.search(rf"\b{heading_pattern}\b", line, re.IGNORECASE):
            for j in range(index, -1, -1):
                page_match = re.search(r'Page\s+(\d+)\s+Text:', lines[j], re.IGNORECASE)
                if page_match:
                    return int(page_match.group(1))
    return None


def resolve_heading_page_number(text, heading, toc_entries=None):
    if not heading:
        return None
    heading_text = str(heading).strip()
    if toc_entries is None:
        toc_entries = extract_toc_with_page_numbers(text)
    for num, title, page_num in toc_entries:
        if title.strip().lower() == heading_text.lower():
            return page_num
        if heading_text.lower() in title.strip().lower() or title.strip().lower() in heading_text.lower():
            return page_num
    return find_heading_page_number(text, heading_text)


def extract_document_headings(text):
    """Extract numbered headings and explicit DOCX headings from extracted text."""
    headings = []
    text = str(text)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    
    for line in lines:
        # Skip lines that are too long
        if len(line) > 120:
            continue
        
        # Skip metadata, page markers, and special content
        if (line.isupper() or line.endswith(":") or 
            "Page" in line or "PDF Metadata" in line or 
            "Total Pages" in line or "TABLE:" in line):
            continue

        # Match explicit heading markers from DOCX extraction
        if line.startswith("Heading:"):
            heading_text = line.replace("Heading:", "", 1).strip()
            if 3 <= len(heading_text) <= 120:
                headings.append(("", heading_text))
            continue
        
        # Match numbered headings at start: "1 Overview", "1.1 Introduction", etc.
        match = re.match(r'^(\d+(?:\.\d+)*)\s+([A-Za-z\s][^.]*?)(?:\s*\.+\s*\d+)?\s*$', line)
        if match:
            num = match.group(1)
            title = match.group(2).strip()
            
            # Clean up any trailing dots or page numbers
            title = re.sub(r'\s*\.+\s*\d*\s*$', '', title).strip()
            
            if 3 <= len(title) <= 120:
                headings.append((num, title))
    
    # Remove duplicates while preserving order
    seen = set()
    deduped = []
    for num, title in headings:
        key = f"{num}:{title}"
        if key not in seen:
            seen.add(key)
            deduped.append((num, title))
    
    return deduped


def extract_toc_with_page_numbers(text):
    """Extract table of contents entries with page numbers from document."""
    toc_entries = []
    text = str(text)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    
    # First, try explicit TOC patterns on full text
    for regex in [
        r'(?m)^\s*(\d+(?:\.\d+)*)\s+(.+?)\s+\.{2,}\s*(\d+)\s*$',
        r'(?m)^\s*(\d+(?:\.\d+)*)\s+(.+?)\s{3,}(\d+)\s*$',
        r'(?m)^\s*(\d+(?:\.\d+)*)\s+(.+?)\s+(\d+)\s*$'
    ]:
        for match in re.finditer(regex, text):
            num = match.group(1)
            title = match.group(2).strip()
            page_num = match.group(3)
            if 3 <= len(title) <= 120 and len(re.findall(r'\d+', title)) <= 2:
                toc_entries.append((num, title, page_num))
        if toc_entries:
            return toc_entries

    # Fallback: build TOC from detected headings and page markers
    headings = extract_document_headings(text)
    if headings:
        for num, title in headings:
            page_num = None
            search_pattern = re.escape(title)
            for i, line in enumerate(lines):
                if title in line or re.search(search_pattern, line, re.IGNORECASE):
                    for j in range(i, max(0, i - 20), -1):
                        page_match = re.search(r'Page\s+(\d+)\s+Text:', lines[j])
                        if page_match:
                            page_num = page_match.group(1)
                            break
                    if page_num:
                        break
            toc_entries.append((num, title, page_num))
    return toc_entries


def build_file_overview(file_name, text):
    text = str(text)
    toc_entries = extract_toc_with_page_numbers(text)
    all_headings = extract_document_headings(text)

    overview_parts = [f"📄 **{file_name}**"]
    
    # Table of Contents section
    overview_parts.append("### Table of Contents")
    if toc_entries:
        overview_parts.append("| Contents | Page No |")
        overview_parts.append("|----------|---------|")
        for num, title, page_num in toc_entries:
            content_str = f"{num} {title}" if num else title
            display_text = f"{content_str} (Page {page_num})" if page_num else content_str
            preview_link = create_preview_link(file_name, highlight_term=title, page_num=page_num)
            anchor_id = create_heading_anchor(title)
            if preview_link:
                page_display = page_num if page_num else "-"
                overview_parts.append(f"| <a href='{preview_link}#{anchor_id}' target='_blank'>{html.escape(display_text)}</a> | {page_display} |")
            else:
                page_display = page_num if page_num else "-"
                overview_parts.append(f"| {html.escape(display_text)} | {page_display} |")
    else:
        overview_parts.append("- No table of contents found with page numbers.")

    # Document Headings section
    overview_parts.append("### Document Headings")
    if all_headings:
        for num, title in all_headings:
            content_str = f"{num} {title}" if num else title
            anchor_id = create_heading_anchor(title)
            page_num = resolve_heading_page_number(text, title, toc_entries)
            preview_link = create_preview_link(file_name, highlight_term=title, page_num=page_num)
            if preview_link:
                overview_parts.append(f"- <a href='{preview_link}#{anchor_id}' target='_blank'>{html.escape(content_str)}</a>")
            else:
                overview_parts.append(f"- {content_str}")
    else:
        overview_parts.append("- No document headings were detected.")

    return "\n".join(overview_parts)


@st.cache_data(show_spinner=False)
def build_highlighted_search_results(file_name, text, query):
    if not query:
        return ""

    pattern = re.compile(re.escape(query), re.IGNORECASE)
    matches = []

    for line_no, raw_line in enumerate(str(text).splitlines(), 1):
        if pattern.search(raw_line):
            escaped_line = html.escape(raw_line)
            highlighted_line = pattern.sub(
                lambda match: f"<mark style='background:#fff3a3; padding:0 2px;'>{html.escape(match.group(0))}</mark>",
                escaped_line
            )
            matches.append(
                f"<div style='margin:0 0 8px 0;'><b>Line {line_no}</b>: {highlighted_line}</div>"
            )

    if not matches:
        return f"<div><b>{html.escape(file_name)}</b><br>No matches found for <code>{html.escape(query)}</code>.</div>"

    return (
        f"<div style='margin-bottom:14px;'>"
        f"<h4 style='margin:0 0 8px 0; color:#1f4f91;'>{html.escape(file_name)} ({len(matches)} matches)</h4>"
        f"{''.join(matches)}"
        f"</div>"
    )


@st.cache_data(show_spinner=False)
def extract_login_name_from_html(file_bytes):
    soup = BeautifulSoup(BytesIO(file_bytes), "html.parser")
    text = soup.get_text(" ", strip=True)
    match = re.search(r'login name[:\s]+(.+?)(version|$)', text, re.IGNORECASE)
    if match:
        name = match.group(1).strip()
        parts = name.split()
        return " ".join(parts[:1])
    return "Not found"


@st.cache_data(show_spinner=False)
def extract_statistics_from_html(file_bytes):
    soup = BeautifulSoup(BytesIO(file_bytes), "html.parser")
    stats = {
        "Executed": 0,
        "Passed": 0,
        "Failed": 0,
        "Inconclusive": 0,
        "Error": 0
    }

    text = soup.get_text(" ", strip=True).lower()
    patterns = {
        "Executed": r'executed test cases[:\s]+(\d+)',
        "Passed": r'passed[:\s]+(\d+)',
        "Failed": r'failed[:\s]+(\d+)',
        "Inconclusive": r'inconclusive[:\s]+(\d+)',
        "Error": r'error[:\s]+(\d+)'
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            stats[key] = int(match.group(1))

    return stats


@st.cache_data(show_spinner=False)
def extract_test_results_grouped_from_html(file_bytes):
    soup = BeautifulSoup(BytesIO(file_bytes), "html.parser")
    results = {}

    group_tables = soup.find_all('table', class_='GroupHeadingTable')

    for group_table in group_tables:
        try:
            rows = group_table.find_all('tr')
            if len(rows) >= 2:
                first_row = rows[0]
                heading = first_row.find('big', class_='Heading3')

                if heading:
                    heading_text = heading.get_text(strip=True)
                    fixture_match = re.search(r'Test Fixture:\s*(.+?)(?:\s|$)', heading_text, re.IGNORECASE)

                    if fixture_match:
                        fixture_name = fixture_match.group(1).strip()
                        second_row = rows[1]
                        overview_table = second_row.find('table', class_='OverviewResultTable')

                        if overview_table:
                            count_cell = overview_table.find('td')
                            if count_cell:
                                try:
                                    count = int(count_cell.get_text(strip=True))

                                    if fixture_name not in results:
                                        results[fixture_name] = {
                                            "name": fixture_name,
                                            "test_cases": [],
                                            "pass": count,
                                            "fail": 0,
                                            "error": 0,
                                            "not executed": 0,
                                            "inconclusive": 0,
                                            "total": count,
                                            "count_cell_class": count_cell.get('class', [''])[0]
                                        }
                                except ValueError:
                                    pass
        except Exception:
            pass

    full_text = soup.get_text("\n", strip=True)
    lines = [l.strip() for l in full_text.split("\n") if l.strip()]

    current_fixture = None

    for i, line in enumerate(lines):
        line_lower = line.lower()

        if "test fixture:" in line_lower:
            fixture_match = re.search(r'Test Fixture:\s*(.+?)(?:\s|$)', line, re.IGNORECASE)
            if fixture_match:
                current_fixture = fixture_match.group(1).strip()
                if current_fixture not in results:
                    results[current_fixture] = {
                        "name": current_fixture,
                        "test_cases": [],
                        "pass": 0,
                        "fail": 0,
                        "error": 0,
                        "not executed": 0,
                        "inconclusive": 0,
                        "total": 0
                    }

        elif re.match(r'^\d+\.\d+', line) and current_fixture:
            verdict_match = re.search(r':\s*(Passed|Failed|Pass|Fail|Error|Not Executed|Inconclusive)\s*$', line,
                                      re.IGNORECASE)

            if verdict_match:
                verdict_word = verdict_match.group(1).lower()

                if "pass" in verdict_word:
                    verdict_type = "Pass"
                    results[current_fixture]["pass"] += 1
                elif "fail" in verdict_word:
                    verdict_type = "Failed"
                    results[current_fixture]["fail"] += 1
                elif "error" in verdict_word:
                    verdict_type = "Error"
                    results[current_fixture]["error"] += 1
                elif "not executed" in verdict_word:
                    verdict_type = "Not Executed"
                    results[current_fixture]["not executed"] += 1
                elif "inconclusive" in verdict_word:
                    verdict_type = "Inconclusive"
                    results[current_fixture]["inconclusive"] += 1
                else:
                    continue

                timestamp = None
                test_step = "Step"
                failure_step_id = ""

                def score_timestamp(candidate):
                    if not candidate:
                        return -1
                    parts = candidate.split('.')
                    if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
                        return -1
                    leading_num = int(parts[0])
                    decimal_places = len(parts[1])
                    decimal_bonus = 10000 if decimal_places >= 3 else (100 if decimal_places == 2 else 0)
                    return decimal_bonus + leading_num

                def find_best_timestamp(text):
                    matches = re.findall(r'\b(\d+\.\d+)\b', text)
                    if not matches:
                        return None
                    return max(matches, key=score_timestamp)

                def find_first_relevant_timestamp(text):
                    for m in re.findall(r'\b(\d+\.\d+)\b', text):
                        if len(m.split('.')[1]) >= 3:
                            return m
                    for m in re.findall(r'\b(\d+\.\d+)\b', text):
                        if len(m.split('.')[1]) >= 2:
                            return m
                    return None

                def consider_timestamp(candidate):
                    nonlocal timestamp
                    if not candidate:
                        return
                    if not timestamp:
                        timestamp = candidate
                        return
                    if len(timestamp.split('.')[1]) >= 3:
                        return
                    if len(candidate.split('.')[1]) > len(timestamp.split('.')[1]):
                        timestamp = candidate
                        return
                    if score_timestamp(candidate) > score_timestamp(timestamp):
                        timestamp = candidate

                same_line_step = re.search(r'(\d+(?:\.\d+)+)\.\s+([^:]+):\s*(failed|fail|error)', line,
                                           re.IGNORECASE)
                if same_line_step:
                    failure_step_id = same_line_step.group(1)
                    action_text = same_line_step.group(2).strip()
                    test_step = action_text
                    consider_timestamp(find_first_relevant_timestamp(line) or find_best_timestamp(line))

                for k in range(i + 1, min(i + 150, len(lines))):
                    next_line = lines[k]

                    if re.match(r'^\d+\.\d+(?:\s|$)', next_line) and k > i + 5:
                        break

                    consider_timestamp(find_first_relevant_timestamp(next_line) or find_best_timestamp(next_line))

                    if verdict_type in ["Failed", "Error"] and not failure_step_id:
                        next_line_lower = next_line.lower()

                        step_match = re.search(r'(\d+(?:\.\d+)+)\.\s+([^:]+):\s*(failed|fail|error)', next_line,
                                               re.IGNORECASE)
                        if step_match:
                            failure_step_id = step_match.group(1)
                            action_text = step_match.group(2).strip()
                            test_step = action_text
                            consider_timestamp(find_best_timestamp(next_line))
                        else:
                            if any(keyword in next_line_lower for keyword in
                                   ["condition", "value", "expected", "actual", "mismatch", "not found",
                                    "exception", "error", "failed to", "failed"]):
                                if not re.match(r'^\d+\.\d+', next_line):
                                    step_num_match = re.match(r'^(\d+(?:\.\d+)*)', next_line.strip())
                                    if step_num_match:
                                        failure_step_id = step_num_match.group(1)
                                        test_step = next_line[:80]

                    if verdict_type == "Pass":
                        next_line_lower = next_line.lower()

                        if "execute" in next_line_lower:
                            match = re.search(r'execute\s+(\w+)', next_line_lower)
                            if match:
                                test_step = match.group(1).capitalize()
                        elif "question" in next_line_lower and "text" in next_line_lower:
                            test_step = "Question/Text"
                        elif "await" in next_line_lower or "wait" in next_line_lower:
                            test_step = "Await Value Match"
                        elif "resume" in next_line_lower:
                            test_step = "Resume"
                        elif "set" in next_line_lower:
                            test_step = "Set"
                        elif "tester" in next_line_lower and "confirmed" in next_line_lower:
                            test_step = "Tester Confirmation"

                if timestamp:
                    results[current_fixture]["test_cases"].append({
                        "timestamp": timestamp,
                        "verdict": verdict_type,
                        "details": test_step
                    })

    for fixture_name in results:
        parsed_count = len(results[fixture_name]["test_cases"])
        initial_count = results[fixture_name].get("total", 0)
        results[fixture_name]["total"] = max(parsed_count, initial_count)

    return results


CREATOR_USERNAME = "Vignesh"
CREATOR_PASSWORD = "Rider@100"

# Login gate:
# This runs before the main app tabs are shown. It keeps the creator/user access
# flow in one place so authentication checks do not have to be repeated per tab.
if not st.session_state.is_authenticated and "preview_token" not in query_params:
    st.subheader("Login")
    login_username = st.text_input("Username")
    login_password = st.text_input("Password", type="password")
    st.info("Note: For users, leave the password field blank (empty).")

    has_read_readme = st.checkbox("I have read the README and want to continue")

    col1, col2 = st.columns([1, 3])
    with col1:
        access_clicked = False
        if has_read_readme:
            access_clicked = st.button("Access", use_container_width=True)
        else:
            st.write("")
    with col2:
        st.write("")

    st.markdown("### Read Me First")
    st.text_area("README", value=README_TEXT, height=360, disabled=True)

    if access_clicked:
        if not has_read_readme:
            st.warning("Please read the README and confirm the checkbox before accessing the app.")
            st.stop()

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
                "timestamp": ist_time
            })
            # Update active users
            active_file = "active_users.json"
            now = datetime.now()
            if os.path.exists(active_file):
                with open(active_file, "r") as f:
                    active_users = json.load(f)
            else:
                active_users = []
            # Clean old entries (>1 hour)
            active_users = [u for u in active_users if
                            datetime.fromisoformat(u["timestamp"]) > now - timedelta(hours=1)]
            # Add current
            active_users.append({"username": cleaned_username, "timestamp": now.isoformat()})
            with open(active_file, "w") as f:
                json.dump(active_users, f)
            st.success("Creator access granted.")
            st.rerun()

        elif len(cleaned_username) > 3 and cleaned_password == "":
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
                "timestamp": ist_time
            })
            # Update active users
            active_file = "active_users.json"
            now = datetime.now()
            if os.path.exists(active_file):
                with open(active_file, "r") as f:
                    active_users = json.load(f)
            else:
                active_users = []
            # Clean old entries (>1 hour)
            active_users = [u for u in active_users if
                            datetime.fromisoformat(u["timestamp"]) > now - timedelta(hours=1)]
            # Add current
            active_users.append({"username": cleaned_username, "timestamp": now.isoformat()})
            with open(active_file, "w") as f:
                json.dump(active_users, f)
            st.success("User access granted.")
            st.rerun()

        else:
            st.error(
                "For creator: use username 'Vignesh' and password 'Rider@100'. For users: username >3 chars, password empty.")

    # st.info("Creator should use Vignesh/Rider@100; others use any login. Creator sees admin features.")
    st.stop()


# Files will be processed on-demand per tab when selected


# -------------------------------
# HELPER CHART FUNCTIONS
# -------------------------------
# Dashboard chart helpers:
# These are used by the Dashboard tab when an XLSX/HTML file is selected and the
# user wants counts shown as bar or pie charts.
def get_column_counts(data, column):
    counts = defaultdict(int)
    for row in data:
        val = row.get(column)
        if val is not None:
            counts[val] += 1
    return dict(counts)


def plot_pie_chart(counts, title):
    labels, values = list(counts.keys()), list(counts.values())
    fig = go.Figure(go.Pie(labels=labels[:50], values=values[:50], textinfo='label+value', textposition='outside'))
    fig.update_layout(title=title, margin=dict(t=80, b=80, l=80, r=80), height=700)
    return fig


def plot_bar_chart(counts, title, horizontal=False):
    labels, values = list(counts.keys()), list(counts.values())
    fig = px.bar(x=values, y=labels, orientation='h', text=values) if horizontal else px.bar(x=labels, y=values,
                                                                                             text=values)
    fig.update_traces(texttemplate='%{text}', textposition='outside', marker_color='skyblue')
    fig.update_layout(title=title, margin=dict(t=80, b=150 if not horizontal else 80), height=700)
    return fig


# -------------------------------
# INLINE MULTI-FILE DIFF (HTML)
# -------------------------------
# Compare tab HTML diff helper:
# Generates the inline visual comparison shown in the Compare tab and also reused
# from Chat when the user asks to compare multiple selected documents.
@st.cache_data(show_spinner=False)
def highlight_multi_file_differences_cached(file_items, comparison_mode="Exact inline word diff", reference_file=None):
    if len(file_items) < 2:
        return "Select at least two files to compare."

    files = [fname for fname, _ in file_items]
    if reference_file is None or reference_file not in files:
        reference_file = files[0]

    css = """
    <style>
        body { font-family: Arial; margin: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid black; padding: 4px; vertical-align: top; white-space: pre-wrap; }
        th { background-color: #f0f0f0; }
        td.line-number { background-color: #f0f0f0; font-weight: bold; text-align: center; }
        .match { background-color: #ccffcc; }
        .mismatch { background-color: #ffcccc; }
        .scrollable { overflow:auto; max-height:800px; }
        p.legend span { display:inline-block; width:20px; height:20px; margin-right:5px; vertical-align:middle; }
    </style>
    """
    html_parts = [
        "<html><head>", css, "</head><body><div class='scrollable'>",
        "<p class='legend'><b>Legend:</b> <span class='match'></span> Matched word, <span class='mismatch'></span> Different or missing word</p>",
        "<table><tr><th>Line #</th>",
        "".join(f"<th>{html.escape(fname)}</th>" for fname in files),
        "</tr>",
    ]

    file_lines = {fname: text.splitlines() for fname, text in file_items}
    max_lines = max(len(lines) for lines in file_lines.values())

    for i in range(max_lines):
        html_parts.append(f"<tr><td class='line-number'>{i + 1}</td>")

        line_word_lists = {}
        ordered_words = []
        word_presence = defaultdict(int)

        for fname in files:
            raw_line = file_lines[fname][i] if i < len(file_lines[fname]) else ""
            words = raw_line.split()
            line_word_lists[fname] = words
            for word in words:
                if word not in ordered_words:
                    ordered_words.append(word)
            for word in set(words):
                word_presence[word] += 1

        reference_words = line_word_lists.get(reference_file, [])

        for fname in files:
            words = line_word_lists[fname]
            if comparison_mode == "Word presence summary":
                highlighted = []
                word_set = set(words)
                for word in ordered_words:
                    escaped_word = html.escape(word)
                    if word in word_set and word_presence[word] == len(files):
                        highlighted.append(f"<span class='match'>{escaped_word}</span>")
                    else:
                        highlighted.append(f"<span class='mismatch'>{escaped_word}</span>")
                cell_html = ' '.join(highlighted) if highlighted else '&nbsp;'
            else:
                highlighted = []
                matcher = SequenceMatcher(None, reference_words, words)
                for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                    if tag == "equal":
                        highlighted.extend(f"<span class='match'>{html.escape(w)}</span>" for w in words[j1:j2])
                    else:
                        highlighted.extend(f"<span class='mismatch'>{html.escape(w)}</span>" for w in words[j1:j2])
                cell_html = ' '.join(highlighted) if highlighted else '&nbsp;'
            html_parts.append(f"<td>{cell_html}</td>")

        html_parts.append("</tr>")

    html_parts.append("</table></div></body></html>")
    return "".join(html_parts)


def highlight_side_by_side_differences_cached(file_items, reference_file=None):
    files = [fname for fname, _ in file_items]
    if len(files) < 2:
        return "Select at least two files to compare."
    if reference_file is None or reference_file not in files:
        reference_file = files[0]

    file_lines = {fname: text.splitlines() for fname, text in file_items}
    max_lines = max(len(lines) for lines in file_lines.values())

    css = """
    <style>
        body { font-family: Arial; margin: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid black; padding: 4px; vertical-align: top; white-space: pre-wrap; }
        th { background-color: #f0f0f0; }
        td.line-number { background-color: #f0f0f0; font-weight: bold; text-align: center; }
        .line-match { background-color: #ccffcc; display: block; width: 100%; }
        .line-mismatch { background-color: #ffcccc; display: block; width: 100%; }
        .scrollable { overflow:auto; max-height:800px; }
        p.legend span { display:inline-block; width:20px; height:20px; margin-right:5px; vertical-align:middle; }
    </style>
    """
    html_parts = [
        "<html><head>", css, "</head><body><div class='scrollable'>",
        "<p class='legend'><b>Legend:</b> <span class='line-match'></span> Same as reference line, <span class='line-mismatch'></span> Different from reference or missing line</p>",
        "<p><b>Reference file:</b> " + html.escape(reference_file) + "</p>",
        "<table><tr><th>Line #</th>",
        "".join(f"<th>{html.escape(fname)}</th>" for fname in files),
        "</tr>",
    ]

    for i in range(max_lines):
        html_parts.append(f"<tr><td class='line-number'>{i + 1}</td>")
        reference_line = file_lines[reference_file][i] if i < len(file_lines[reference_file]) else ""
        for fname in files:
            line_text = file_lines[fname][i] if i < len(file_lines[fname]) else ""
            if line_text == reference_line and line_text != "":
                cell_html = f"<span class='line-match'>{html.escape(line_text)}</span>"
            elif line_text == reference_line == "":
                cell_html = "&nbsp;"
            else:
                cell_html = f"<span class='line-mismatch'>{html.escape(line_text)}</span>"
            html_parts.append(f"<td>{cell_html}</td>")
        html_parts.append("</tr>")

    html_parts.append("</table></div></body></html>")
    return "".join(html_parts)


def highlight_multi_file_differences(file_texts, comparison_mode="Exact inline word diff", reference_file=None):
    if comparison_mode == "Side-by-side line diff":
        return highlight_side_by_side_differences_cached(
            tuple((fname, str(text)) for fname, text in file_texts.items()),
            reference_file=reference_file
        )
    return highlight_multi_file_differences_cached(
        tuple((fname, str(text)) for fname, text in file_texts.items()),
        comparison_mode=comparison_mode,
        reference_file=reference_file
    )


# -------------------------------
# COMPARE EXCEL HIGHLIGHT
# -------------------------------
# Compare tab Excel export helper:
# Builds the downloadable workbook used in the Compare tab so users can inspect
# mismatches outside the Streamlit UI.
@st.cache_data(show_spinner=False)
def generate_word_level_comparison_excel_cached(file_items):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comparison"

    files = [fname for fname, _ in file_items]
    file_texts = {fname: text for fname, text in file_items}
    ws.append(["Line #"] + files)
    file_lines = {f: [l.split() for l in t.splitlines()] for f, t in file_texts.items()}

    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

    max_lines = max(len(l) for l in file_lines.values())

    for i in range(max_lines):
        max_words = max(len(file_lines[f][i]) if i < len(file_lines[f]) else 0 for f in files)
        for w_idx in range(max_words):
            row_values = [i + 1 if w_idx == 0 else ""]
            for f in files:
                line_words = file_lines[f][i] if i < len(file_lines[f]) else []
                word = line_words[w_idx] if w_idx < len(line_words) else ""
                row_values.append(word)
            ws.append(row_values)

            # Highlight exact matches in green and missing/mismatched content in red
            all_words_set = set()
            for f in files:
                if i < len(file_lines[f]):
                    all_words_set.update(file_lines[f][i])
            for col_idx, f in enumerate(files, start=2):
                cell = ws.cell(row=ws.max_row, column=col_idx)
                line_words = file_lines[f][i] if i < len(file_lines[f]) else []
                if w_idx >= len(line_words):
                    cell.fill = red_fill
                elif all(word == line_words[w_idx] for other_file in files
                         for word in ([file_lines[other_file][i][w_idx]]
                                      if i < len(file_lines[other_file]) and w_idx < len(file_lines[other_file][i])
                                      else ["__missing__"])):
                    cell.fill = green_fill
                else:
                    cell.fill = red_fill

    excel_io = BytesIO()
    wb.save(excel_io)
    return excel_io.getvalue()


def generate_word_level_comparison_excel(file_texts):
    excel_io = BytesIO(generate_word_level_comparison_excel_cached(tuple((fname, str(text)) for fname, text in file_texts.items())))
    excel_io.seek(0)
    return excel_io


# -------------------------------
# CAPL Complier
# -------------------------------
# CAPL analyzer helpers:
# These functions are used only by the CAPL tab for syntax checking, issue
# listing, and highlighted code rendering inside the CAPL editor/viewer panels.
@st.cache_data(show_spinner=False)
def analyze_capl_code_with_suggestions_cached(code):
    issues = []

    brace_stack = []
    declared_vars = []
    used_vars = []

    lines = code.splitlines()

    for i, line in enumerate(lines, 1):
        stripped = line.strip()

        if not stripped or stripped.startswith("//"):
            continue  # Skip empty lines/comments

        # Track braces
        for c in stripped:
            if c == "{":
                brace_stack.append(i)
            elif c == "}":
                if brace_stack:
                    brace_stack.pop()
                else:
                    issues.append({
                        "line": i,
                        "error": "Unmatched closing brace",
                        "suggestion": "Remove or match with an opening '{'"
                    })

        # Detect variable declarations
        var_match = re.match(r'\b(int|float|byte|char|mstimer|timer|enum)\b\s+(\w+)', stripped)
        if var_match:
            declared_vars.append(var_match.group(2))

        # Track all used variable names
        used_vars += re.findall(r'\b([a-zA-Z_]\w*)\b', stripped)

        # Check for case sensitivity in keywords
        if re.search(r'\b(If|Else|For|While|Switch|Case|Return|On|Variables|Includes|Enum|Mstimer|Timer)\b', stripped):
            issues.append({
                "line": i,
                "error": "CAPL keywords should be lowercase",
                "suggestion": "Use lowercase keywords like 'if', 'else', 'on', etc."
            })

        # Check for incomplete if conditions
        if re.match(r'^\s*(if|else if)\s*\(', stripped) and not re.search(r'\)\s*(\{)?\s*$', stripped):
            issues.append({
                "line": i,
                "error": "Incomplete if condition",
                "suggestion": "Add closing parenthesis ')' and possibly opening brace '{'"
            })

        # Check for missing opening brace after control statements
        if re.match(r'^\s*(if|else if|else|for|while|switch)\b', stripped) and not stripped.endswith(
                '{') and not re.search(r'\)\s*\{', stripped):
            # Check if next line starts with '{'
            if i < len(lines) and not lines[i].strip().startswith('{'):
                issues.append({
                    "line": i,
                    "error": "Missing opening brace after control statement",
                    "suggestion": "Add '{' after the condition or on the next line"
                })

        # Detect missing semicolon
        if not stripped.endswith(";") and not stripped.endswith("{") and not stripped.endswith("}"):
            if not re.match(r'^(on|variables|includes|enum|mstimer|timer|if|else|switch|case|for|while|return)\b',
                            stripped):
                issues.append({
                    "line": i,
                    "error": "Missing semicolon",
                    "suggestion": "Add ';' at the end of this line"
                })

    # Check unmatched opening braces
    for open_line in brace_stack:
        issues.append({
            "line": open_line,
            "error": "Unmatched opening brace",
            "suggestion": "Add closing '}' to match this '{'"
        })

    # Check for 'on message' presence
    if "on message" not in code.lower():
        issues.append({
            "line": None,
            "error": "No 'on message' handler found",
            "suggestion": "Add an 'on message' event handler as required"
        })

    # Check for unused declared variables
    for var in declared_vars:
        if var not in used_vars:
            issues.append({
                "line": None,
                "error": f"Unused variable: {var}",
                "suggestion": "Consider removing this variable or using it in the code"
            })

    # Detect undeclared variables starting with PT4_ or $PT4_ used in code
    for i, line in enumerate(lines, 1):
        pt4_vars = re.findall(r'\b(PT4_[a-zA-Z_]\w*|\$PT4_[a-zA-Z_]\w*)\b', line)
        for var in pt4_vars:
            if var not in declared_vars and not var.startswith("$"):
                issues.append({
                    "line": i,
                    "error": f"Undeclared variable used: {var}",
                    "suggestion": f"Declare '{var}' in the variables section before using it"
                })

    return issues


def analyze_capl_code_with_suggestions(code):
    return analyze_capl_code_with_suggestions_cached(code)


# -------------------------------
# CAPL CODE DETECTION
# -------------------------------
def is_capl_code(text):
    """Check if the given text contains CAPL-specific keywords or syntax."""
    capl_keywords = [
        "on message", "variables", "includes", "mstimer", "timer", "byte", "char", "int", "float",
        "enum", "if", "else", "switch", "case", "for", "while", "return", "write", "output",
        "setTimer", "cancelTimer", "getTimer", "putValue", "getValue", "testcase", "teststep"
    ]
    text_lower = text.lower()
    return any(keyword in text_lower for keyword in capl_keywords)


@st.cache_data(show_spinner=False)
def render_capl_code_with_highlights_cached(code, issues_key):
    """Render CAPL code with IDE-like line highlighting for detected issues."""
    issues = [
        {"line": line, "error": error, "suggestion": suggestion}
        for line, error, suggestion in issues_key
    ]
    issue_lines = defaultdict(list)

    for issue in issues:
        line_no = issue.get("line")
        if isinstance(line_no, int):
            issue_lines[line_no].append(issue.get("error", "Issue detected"))

    code_lines = code.splitlines() or [""]
    rendered_lines = []

    for line_no, line in enumerate(code_lines, 1):
        escaped_line = html.escape(line) if line else "&nbsp;"
        line_classes = ["capl-line"]
        if line_no in issue_lines:
            line_classes.append("capl-line-error")

        tooltip = " | ".join(issue_lines[line_no]) if line_no in issue_lines else ""
        title_attr = f' title="{html.escape(tooltip)}"' if tooltip else ""

        rendered_lines.append(
            f"<div class=\"{' '.join(line_classes)}\"{title_attr}>"
            f"<span class=\"capl-gutter\">{line_no:>4}</span>"
            f"<span class=\"capl-code-text\">{escaped_line}</span>"
            f"</div>"
        )

    code_html = """
    <style>
        .capl-code-block {
            background: #0f172a;
            border: 1px solid #cbd5e1;
            border-radius: 10px;
            font-family: Consolas, "Courier New", monospace;
            font-size: 14px;
            line-height: 1.5;
            max-height: 420px;
            overflow: auto;
            padding: 12px 0;
        }
        .capl-line {
            color: #e2e8f0;
            display: flex;
            white-space: pre;
        }
        .capl-line-error {
            background: rgba(239, 68, 68, 0.22);
            border-left: 4px solid #ef4444;
        }
        .capl-gutter {
            color: #94a3b8;
            display: inline-block;
            min-width: 52px;
            padding: 0 12px;
            text-align: right;
            user-select: none;
        }
        .capl-code-text {
            display: inline-block;
            padding: 0 16px 0 0;
            width: 100%;
        }
    </style>
    """
    return code_html + f"<div class='capl-code-block'>{''.join(rendered_lines)}</div>"


def render_capl_code_with_highlights(code, issues=None):
    issues_key = tuple(
        (
            issue.get("line"),
            issue.get("error", "Issue detected"),
            issue.get("suggestion", "")
        )
        for issue in (issues or [])
    )
    return render_capl_code_with_highlights_cached(code, issues_key)


def render_capl_issue_table(issues):
    if not issues:
        st.success("✅ No issues detected!")
        return

    df_issues = pd.DataFrame(issues).fillna("-")
    st.dataframe(df_issues, use_container_width=True, hide_index=True)


def get_combined_vector_store(file_names):
    ensure_files_processed(file_names)
    selection_key = get_selection_signature(file_names)
    if selection_key not in st.session_state.vector_stores:
        combined_text = "\n".join(st.session_state.file_texts.get(file_name, "") for file_name in file_names)
        st.session_state.vector_stores[selection_key] = create_vector_store(combined_text)
    return st.session_state.vector_stores[selection_key]


# Shared UI helpers:
# These small functions are reused across multiple tabs to show the current
# sidebar selection, tab-level file context, and floating helper popups.
def show_current_sidebar_selection():
    selected = st.session_state.get("selected_files", [])
    if selected:
        st.info("Sidebar selected files: " + ", ".join(selected))
    else:
        st.info("No sidebar files selected yet. Upload and select files from the sidebar first.")


def render_file_context_card(title, available_files, active_files=None):
    active_files = active_files or []
    chips_html = "".join(
        f"<span class='file-chip'>{html.escape(file_name)}</span>"
        for file_name in active_files[:12]
    )
    if len(active_files) > 12:
        chips_html += f"<span class='file-chip'>+{len(active_files) - 12} more</span>"

    st.markdown(
        f"""
        <div class="app-card">
            <h4>{html.escape(title)}</h4>
            <p class="app-muted">Available from sidebar: {len(available_files)} file(s)</p>
            <p class="app-muted">Selected in this tab: {len(active_files)} file(s)</p>
            <div class="file-chip-wrap">
                {chips_html if chips_html else "<span class='file-chip'>No tab files selected yet</span>"}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


def _help_state_key(tab_name):
    return f"show_help_popup_{tab_name}"


def ensure_help_popup_state(tab_name):
    key = _help_state_key(tab_name)
    if key not in st.session_state:
        st.session_state[key] = False
    return key


# Define keywords for each tab
tab_keywords = {
    "chat": ["overview", "summary", "count", "find", "analyze", "explain", "details", "questions"],
    "dashboard": ["statistics", "trends", "charts", "metrics", "data", "visualize", "insights"],
    "compare": ["differences", "compare", "changes", "merge", "diff", "side-by-side", "inline"],
    "capl": ["syntax", "variables", "functions", "errors", "debug", "code", "fix", "suggestions"]
}


def show_help_popup(tab_name, selected_files):
    state_key = ensure_help_popup_state(tab_name)
    st.checkbox("Show query helper popup", key=state_key)

    if not st.session_state[state_key]:
        return

    keywords = tab_keywords.get(tab_name.lower(), [])

    if not selected_files:
        st.markdown(
            f"""
            <div style='position:fixed; bottom:14px; right:14px; width:340px; padding:14px; background:#ffffff; border:1px solid #1f4f91; border-radius:12px; box-shadow:0 8px 24px rgba(0,0,0,0.24); z-index:9999;'>
                <h4 style='margin:0 0 6px 0; font-size:15px; color:#1f4f91;'>📘 {tab_name.capitalize()} Query Keywords</h4>
                <p style='margin:0; font-size:13px; color:#253659;'>Select a document first to see targeted query guidance.</p>
                <p style='margin:5px 0 0 0; font-size:13px; color:#253659;'>Suggested keywords: {', '.join(keywords)}</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        return

    selected_types = {os.path.splitext(f)[1].lower() for f in selected_files}
    if tab_name == "chat":
        extra = "Try count('Approved'), find('Error'), summarize(), or specific queries like 'How many tests passed?'"
    elif tab_name == "dashboard":
        extra = "If *.xlsx selected, ask for columns and trend counts; *.html can be parsed for test verdicts."
    elif tab_name == "compare":
        extra = "Choose inline or side-by-side diff, select a reference file, and download results for PDF, DOCX, PPTX, XLSX, HTML, TXT, CAN, and CAPL files."
    elif tab_name == "capl":
        extra = "For CAPL code, use 'find missing semicolon' or 'show unused variables' style queries."
    else:
        extra = "Use text queries about the document content."

    type_hint = ""
    if ".xlsx" in selected_types:
        type_hint = "For spreadsheets, reference specific columns like 'column 2' or 'total defects'."
    elif ".html" in selected_types or ".htm" in selected_types:
        type_hint = "For HTML, ask about login name, pass/fail stats, and fixture totals."
    elif any(ext in selected_types for ext in [".txt", ".pdf", ".docx", ".pptx"]):
        type_hint = "For full text files, use keywords and exact phrase search in your query."
    elif any(ext in selected_types for ext in [".capl", ".can"]):
        type_hint = "CAPL scripts are analyzed for syntax issues; ask for code fixes or suggestions."

    st.markdown(
        f"""
        <div style='position:fixed; bottom:14px; right:14px; width:340px; padding:14px; background:#ffffff; border:1px solid #1f4f91; border-radius:12px; box-shadow:0 8px 24px rgba(0,0,0,0.24); z-index:9999;'>
            <h4 style='margin:0 0 6px 0; font-size:15px; color:#1f4f91;'>📘 {tab_name.capitalize()} Query Keywords</h4>
            <p style='margin:0 0 8px 0; font-size:13px; color:#253659;'>Support for selected file types: {', '.join(sorted(selected_types))}</p>
            <p style='margin:0 0 8px 0; font-size:13px; color:#253659;'>Suggested keywords: {', '.join(keywords)}</p>
            <p style='margin:0; font-size:12px; color:#3c4f7e;'>{extra}</p>
            <p style='margin:5px 0 0 0; font-size:12px; color:#3c4f7e;'>{type_hint}</p>
        </div>
        """,
        unsafe_allow_html=True
    )


# -------------------------------
# TABS
# -------------------------------
# Main application tabs:
# Each block below owns one visible area of the app. If you want to change a
# feature, start in the matching tab block and then follow the helper comments above.

# Session-backed main navigation:
# 1. Premium "Soft-Glow" Navigation CSS
st.markdown("""
    <style>
    /* Hide default radio UI */
    div[role="radiogroup"] > label > div:first-child { display: none !important; }
    div[role="radiogroup"] { gap: 12px; display: flex; }

    /* Base Pill Styling */
    div[role="radiogroup"] > label {
        background-color: rgba(128, 128, 128, 0.08) !important;
        padding: 8px 22px !important;
        border-radius: 50px !important;
        border: 1px solid rgba(128, 128, 128, 0.1) !important;
        display: flex !important;
        align-items: center !important;
        height: 42px;
        font-weight: 500;
        transition: all 0.4s cubic-bezier(0.23, 1, 0.32, 1);
    }

    /* Active State - Deep Electric Blue */
    div[role="radiogroup"] > label[data-checked="true"] {
        background-color: #1E88E5 !important;
        color: white !important;
        box-shadow: 0 4px 15px rgba(30, 136, 229, 0.4);
    }

    /* Dot Base Style - Soft Glow */
    div[role="radiogroup"] > label::after {
        content: '';
        margin-left: 12px;
        width: 6px;
        height: 6px;
        border-radius: 50%;
        filter: blur(0.4px); /* Softens the edges */
    }

    /* White Dot for Active Tab */
    div[role="radiogroup"] > label[data-checked="true"]::after {
        background-color: white !important;
        box-shadow: 0 0 8px #ffffff;
    }

    /* --- UPDATED PREMIUM COLORS --- */
    
    /* 1. CHAT - Soft Ice Blue (Instead of bright cyan) */
    div[role="radiogroup"] > label:nth-child(1):not([data-checked="true"])::after { 
        background-color: #81D4FA; 
        box-shadow: 0 0 6px #81D4FA;
    } 

    /* 2. DASHBOARD - Emerald Mint (Softer than standard Green) */
    div[role="radiogroup"] > label:nth-child(2):not([data-checked="true"])::after { 
        background-color: #66BB6A; 
        box-shadow: 0 0 6px #66BB6A;
    } 

    /* 3. COMPARE - Muted Amber (Less "Caution" yellow) */
    div[role="radiogroup"] > label:nth-child(3):not([data-checked="true"])::after { 
        background-color: #FFB74D; 
        box-shadow: 0 0 6px #FFB74D;
    } 

    /* 4. CAPL - Royal Orchid (Sophisticated Purple) */
    div[role="radiogroup"] > label:nth-child(4):not([data-checked="true"])::after { 
        background-color: #BA68C8; 
        box-shadow: 0 0 6px #BA68C8;
    }
    </style>
""", unsafe_allow_html=True)

# 2. Your Tab Logic
main_tab_options = ["💬 Chat", "📊 Dashboard", "📂 Compare", "🧠 CAPL"]
active_main_tab = st.radio("Open Section", main_tab_options, horizontal=True, key="active_main_tab", label_visibility="collapsed")

# -------------------------------
# TAB 1: CHAT
# -------------------------------
# Chat tab:
# Handles per-tab file selection, direct commands like summarize/find/count,
# semantic Q&A via vector search, and chat-specific download assets.
if active_main_tab == "💬 Chat":
    chat_header_col, chat_reset_col = st.columns([8, 1])
    with chat_header_col:
        st.subheader("Chat with Selected Documents")
    with chat_reset_col:
        if st.button("🔄 Reset", key="reset_chat_selection", use_container_width=True):
            st.session_state.chat_file_selection = []
            st.session_state.chat_summary_downloads = {"images": [], "tables": []}
            st.session_state.messages = []
            st.success("✅ Chat reset!")
            st.rerun()

    st.info(
        "Choose files in the sidebar to make them available here. Then select only the files you want for Chat in this tab.")
    show_current_sidebar_selection()
    render_file_context_card("Chat File Context", st.session_state.selected_files, st.session_state.chat_file_selection)

    show_help_popup('chat', st.session_state.selected_files)

    if st.session_state.selected_files:
        st.session_state.chat_file_selection = [
            file_name for file_name in st.session_state.chat_file_selection
            if file_name in st.session_state.selected_files
        ]
        chat_files = st.multiselect("Choose file(s) for Chat", options=st.session_state.selected_files,
                                    default=st.session_state.chat_file_selection, key="chat_file_selection")
        if not chat_files:
            st.info("Choose one or more files in this tab to start chatting.")
        else:
            with st.spinner("Loading selected files..."):
                ensure_files_processed(chat_files)
            combined_text = "\n".join([st.session_state.file_texts.get(f, "") for f in chat_files])
    
            user_input = st.chat_input("Ask something... (type 'clear' to reset chat). Try: 'summarize', 'find:', 'count:', or 'overview'")
            if user_input:
                if user_input.strip().lower() == "clear":
                    st.session_state.messages = []
                    st.session_state.chat_summary_downloads = {"images": [], "tables": []}
                    st.success("✅ Chat cleared!")
                else:
                    st.session_state.messages.append({"role": "user", "content": user_input})
                    with st.spinner("Processing your request..."):
                        st.session_state.chat_summary_downloads = {"images": [], "tables": []}
                        user_input_lower = user_input.lower()
                        # Word count queries
                        if any(t in user_input_lower for t in ["how many", "count", "number of", "occurrences"]):
                            match = re.search(r"'(.*?)'|\"(.*?)\"", user_input)
                            if match:
                                word = match.group(1) or match.group(2)
                                count = len(
                                    re.findall(rf'(?<![\w-]){re.escape(word)}(?![\w-])', combined_text, re.IGNORECASE))
                                response = f"🔢 The word/phrase '{word}' appears {count} times in the selected documents."
                            else:
                                response = "⚠️ Specify the word/phrase in quotes. Example: count('keyword') or count(\"keyword\")"
                        elif any(term in user_input_lower for term in ["find", "search", "locate"]) or "highlight" in user_input_lower:
                            match = re.search(r"'(.*?)'|\"(.*?)\"", user_input)
                            if match:
                                query = match.group(1) or match.group(2)
                                response_blocks = []
                                for f in chat_files:
                                    file_text = st.session_state.file_texts.get(f, "")
                                    response_blocks.append(build_highlighted_search_results(f, file_text, query))
                                response = "".join(response_blocks)
                            else:
                                response = "⚠️ Specify the search word or phrase in quotes. Example: find('keyword') or search(\"keyword\")"
                        elif "overview" in user_input_lower:
                            response_lines = []
                            for f in chat_files:
                                file_text = st.session_state.file_texts.get(f, "")
                                if file_text.strip():
                                    toc_entries = extract_toc_with_page_numbers(file_text)
                                    all_headings = extract_document_headings(file_text)
                                    if all_headings:
                                        response_lines.append(f"📄 **{f}**")
                                        response_lines.append("### Document Headings")
                                        for num, title in all_headings:
                                            content_str = f"{num} {title}" if num else title
                                            page_num = resolve_heading_page_number(file_text, title, toc_entries)
                                            display_text = f"{content_str} (Page {page_num})" if page_num else content_str
                                            preview_link = create_preview_link(f, highlight_term=title, page_num=page_num)
                                            if preview_link:
                                                anchor_id = create_heading_anchor(title)
                                                response_lines.append(f"- <a href='{preview_link}#{anchor_id}' target='_blank'>{html.escape(display_text)}</a>")
                                            else:
                                                response_lines.append(f"- {html.escape(display_text)}")
                                    else:
                                        response_lines.append(f"📄 **{f}**\n\nNo document headings were detected.")
                                else:
                                    response_lines.append(f"📄 **{f}**\n\nNo readable content found in this document.")
                            response = "\n\n".join(response_lines)
                        elif any(term in user_input_lower for term in ["analyze", "summary", "summarize", "summarise"]):
                            result = []
                            summary_image_downloads = []
                            summary_table_downloads = []
                            for f in chat_files:
                                file_text = st.session_state.file_texts.get(f, "")
                                file_entry = get_uploaded_file_entry(f)
                                if file_text.strip():
                                    file_bytes = file_entry["bytes"] if file_entry else b""
                                    result.append(build_detailed_document_summary(f, file_bytes, file_text))
                                    summary_assets = build_summary_download_assets(f, file_bytes)
                                    summary_image_downloads.extend(summary_assets.get("images", []))
                                    summary_table_downloads.extend(summary_assets.get("tables", []))
                                else:
                                    result.append(f"📄 **{f}**\n\nNo readable content found in this document.")

                            st.session_state.chat_summary_downloads = {
                                "images": summary_image_downloads,
                                "tables": summary_table_downloads
                            }
                            response = "\n\n---\n\n".join(result)
                        elif "compare" in user_input_lower:
                            if len(chat_files) >= 2:
                                selected_texts = {f: st.session_state.file_texts[f] for f in chat_files}
                                response = highlight_multi_file_differences(selected_texts)
                            else:
                                response = "⚠️ Please select at least 2 files to compare documents."
                        else:
                            combined_vs = get_combined_vector_store(chat_files)
                            retriever = combined_vs.as_retriever(search_kwargs={"k": 3})
                            llm = load_llm()
                            prompt = ChatPromptTemplate.from_messages([
                                ("system",
                                 "You are an intelligent document assistant. Answer ONLY using context from the provided documents.\nIf information is not found in the documents, say 'This information is not available in the uploaded documents.'\nContext:\n{context}"),
                                ("human", "{question}")
                            ])
                            chain = None
                            if llm is not None:
                                try:
                                    chain = ({"context": retriever | (lambda x: '\n'.join(x)),
                                              "question": RunnablePassthrough()} | prompt | llm)
                                except Exception as e:
                                    st.warning(f"Could not create LLM chain: {e}")
                                    chain = None

                            if chain is not None:
                                response = str(chain.invoke(user_input))
                            else:
                                response = "⚠️ AI model is unavailable. Use direct extraction questions such as 'count(\"keyword\")', 'find(\"phrase\")', 'summarize', or 'overview'."
                        st.session_state.messages.append({"role": "assistant", "content": response})

        for msg in st.session_state.messages:
            role = "🧑" if msg["role"] == "user" else "🤖"
            st.markdown(f"{role} {msg['content']}", unsafe_allow_html=True)

        render_chat_summary_downloads()
    else:
        st.info("Select files from the sidebar to start chatting.")

# -------------------------------
# TAB 2: DASHBOARD
# -------------------------------
# Dashboard tab:
# Focused on structured HTML/XLSX analysis, charts, login/stat extraction, and
# grouped test fixture reporting for uploaded report files.
if active_main_tab == "📊 Dashboard":
    dashboard_header_col, dashboard_reset_col = st.columns([8, 1])
    with dashboard_header_col:
        st.subheader("Dashboard")
    with dashboard_reset_col:
        if st.button("🔄 Reset", key="reset_dashboard_selection", use_container_width=True):
            st.session_state.file_dropdown = "--Select File--"
            st.rerun()

    show_current_sidebar_selection()
    show_help_popup('dashboard', [
        f for f in st.session_state.selected_files
        if f.lower().endswith((".html", ".htm", ".xlsx"))
    ])

    # Filter selected files for dashboard-compatible formats
    dashboard_files = [
        f for f in st.session_state.selected_files
        if f.lower().endswith((".html", ".htm", ".xlsx"))
    ]
    active_dashboard_files = [] if st.session_state.file_dropdown == "--Select File--" else [st.session_state.file_dropdown]
    render_file_context_card("Dashboard File Context", dashboard_files, active_dashboard_files)


    def clean_text(x):
        return re.sub(r"\s+", " ", x).strip().lower()


    def safe_int(x):
        nums = re.findall(r"\d+", str(x))
        return int(nums[0]) if nums else 0


    def sort_key(x):
        return [int(i) for i in x.split(".")]


    def format_label(x):
        level = x.count(".")
        return "   " * level + x


    def extract_timestamp_from_line(line):
        """Extract timestamp or identifier from a verdict line"""
        # Try to extract decimal number first (e.g., 35.877473)
        decimal_match = re.search(r'\b(\d+\.\d+)\b', line)
        if decimal_match:
            return decimal_match.group(1)

        # Try to extract time pattern (HH:MM:SS or HH:MM)
        time_match = re.search(r'\b(\d{1,2}:\d{2}:\d{2})\b', line)
        if time_match:
            return time_match.group(1)

        # Try to extract date-time pattern
        datetime_match = re.search(r'\d{4}-\d{2}-\d{2}\s+\d{1,2}:\d{2}:\d{2}', line)
        if datetime_match:
            return datetime_match.group(0)

        # Try to extract test case ID (e.g., "1.", "Test Case 1", "TC_001")
        id_match = re.search(r'(?:test\s+case|tc)[\s_-]*(\d+)', line, re.IGNORECASE)
        if id_match:
            return f"TC {id_match.group(1)}"

        # Extract first number or identifier
        num_match = re.search(r'^\d+', line)
        if num_match:
            return num_match.group(0)

        # Extract the first meaningful text before colon or special char
        first_word = re.match(r'^([a-zA-Z0-9_\-\.]+)', line)
        if first_word:
            return first_word.group(1)

        # Fallback to full line (truncated)
        return line[:30]


    def extract_test_results_grouped(soup):
        """
        Extract test results from HTML grouped by test fixtures.
        Dynamically discovers fixtures from HTML structure and their numerical counts.
        Only counts the main test case verdict, not sub-step verdicts.
        Captures detailed failure information for failed test cases (step ID, action, timestamp).
        """
        results = {}

        # First, extract fixture names and counts from GroupHeadingTable structures
        group_tables = soup.find_all('table', class_='GroupHeadingTable')

        for group_table in group_tables:
            try:
                rows = group_table.find_all('tr')
                if len(rows) >= 2:
                    # First row contains fixture name in Heading3
                    first_row = rows[0]
                    heading = first_row.find('big', class_='Heading3')

                    if heading:
                        heading_text = heading.get_text(strip=True)
                        fixture_match = re.search(r'Test Fixture:\s*(.+?)(?:\s|$)', heading_text, re.IGNORECASE)

                        if fixture_match:
                            fixture_name = fixture_match.group(1).strip()

                            # Second row contains the count in OverviewResultTable
                            second_row = rows[1]
                            overview_table = second_row.find('table', class_='OverviewResultTable')

                            if overview_table:
                                count_cell = overview_table.find('td')
                                if count_cell:
                                    try:
                                        count = int(count_cell.get_text(strip=True))

                                        if fixture_name not in results:
                                            results[fixture_name] = {
                                                "name": fixture_name,
                                                "test_cases": [],
                                                "pass": count,  # All are passed by default from green cells
                                                "fail": 0,
                                                "error": 0,
                                                "not executed": 0,
                                                "inconclusive": 0,
                                                "total": count,
                                                "count_cell_class": count_cell.get('class', [''])[0]
                                            }
                                    except ValueError:
                                        pass
            except Exception as e:
                pass

        # Parse full text for verdict distribution
        full_text = soup.get_text("\n", strip=True)
        lines = [l.strip() for l in full_text.split("\n") if l.strip()]

        current_fixture = None

        for i, line in enumerate(lines):
            line_lower = line.lower()

            # Check for fixture marker
            if "test fixture:" in line_lower:
                fixture_match = re.search(r'Test Fixture:\s*(.+?)(?:\s|$)', line, re.IGNORECASE)
                if fixture_match:
                    current_fixture = fixture_match.group(1).strip()
                    if current_fixture not in results:
                        results[current_fixture] = {
                            "name": current_fixture,
                            "test_cases": [],
                            "pass": 0,
                            "fail": 0,
                            "error": 0,
                            "not executed": 0,
                            "inconclusive": 0,
                            "total": 0
                        }

            # Check for test case number with verdict on SAME line (e.g., "1.2.1 ...description...: Passed")
            # Only match if line contains both test number and a verdict
            elif re.match(r'^\d+\.\d+', line) and current_fixture:
                verdict_match = re.search(r':\s*(Passed|Failed|Pass|Fail|Error|Not Executed|Inconclusive)\s*$', line,
                                          re.IGNORECASE)

                if verdict_match:
                    verdict_word = verdict_match.group(1).lower()

                    # Normalize verdict and increment counter
                    if "pass" in verdict_word:
                        verdict_type = "Pass"
                        results[current_fixture]["pass"] += 1
                    elif "fail" in verdict_word:
                        verdict_type = "Failed"
                        results[current_fixture]["fail"] += 1
                    elif "error" in verdict_word:
                        verdict_type = "Error"
                        results[current_fixture]["error"] += 1
                    elif "not executed" in verdict_word:
                        verdict_type = "Not Executed"
                        results[current_fixture]["not executed"] += 1
                    elif "inconclusive" in verdict_word:
                        verdict_type = "Inconclusive"
                        results[current_fixture]["inconclusive"] += 1
                    else:
                        continue

                    # Now look forward for timestamp and failure details
                    timestamp = None
                    test_step = "Step"
                    failure_step_id = ""

                    def score_timestamp(candidate):
                        if not candidate:
                            return -1
                        parts = candidate.split('.')
                        if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
                            return -1
                        leading_num = int(parts[0])
                        decimal_places = len(parts[1])
                        decimal_bonus = 10000 if decimal_places >= 3 else (100 if decimal_places == 2 else 0)
                        return decimal_bonus + leading_num

                    def find_best_timestamp(text):
                        # Find all decimal numbers
                        matches = re.findall(r'\b(\d+\.\d+)\b', text)
                        if not matches:
                            return None

                        return max(matches, key=score_timestamp)

                    def find_first_relevant_timestamp(text):
                        # Prefer first high-precision timestamp (3+ decimals) in text
                        for m in re.findall(r'\b(\d+\.\d+)\b', text):
                            if len(m.split('.')[1]) >= 3:
                                return m
                        # fallback to lower precision (2+ decimals)
                        for m in re.findall(r'\b(\d+\.\d+)\b', text):
                            if len(m.split('.')[1]) >= 2:
                                return m
                        return None

                    def consider_timestamp(candidate):
                        nonlocal timestamp
                        if not candidate:
                            return
                        # Prefer the first nearby relevant timestamp (prefer earlier block context)
                        if not timestamp:
                            timestamp = candidate
                            return
                        # if already have a high precision timestamp, keep it
                        if len(timestamp.split('.')[1]) >= 3:
                            return
                        # if candidate is higher precision than existing timestamp, replace
                        if len(candidate.split('.')[1]) > len(timestamp.split('.')[1]):
                            timestamp = candidate
                            return
                        # as fallback, use original scoring strategy
                        if score_timestamp(candidate) > score_timestamp(timestamp):
                            timestamp = candidate

                    same_line_step = re.search(r'(\d+(?:\.\d+)+)\.\s+([^:]+):\s*(failed|fail|error)', line,
                                               re.IGNORECASE)
                    if same_line_step:
                        failure_step_id = same_line_step.group(1)
                        action_text = same_line_step.group(2).strip()
                        test_step = action_text  # prefer just action text as details
                        # timestamp candidate in same line: prefer first relevant and then best
                        consider_timestamp(find_first_relevant_timestamp(line) or find_best_timestamp(line))

                    for k in range(i + 1, min(i + 150, len(lines))):
                        next_line = lines[k]

                        # Stop if we hit next test case
                        if re.match(r'^\d+\.\d+(?:\s|$)', next_line) and k > i + 5:
                            break

                        # Look for timestamp (decimal number) - general case
                        consider_timestamp(find_first_relevant_timestamp(next_line) or find_best_timestamp(next_line))

                        # For Failed/Error verdicts, look for failed step details
                        if verdict_type in ["Failed", "Error"] and not failure_step_id:
                            next_line_lower = next_line.lower()

                            # Look for step identifier with action (e.g., "10.1.6.9.4. Await Value Match: Failed")
                            step_match = re.search(r'(\d+(?:\.\d+)+)\.\s+([^:]+):\s*(failed|fail|error)', next_line,
                                                   re.IGNORECASE)
                            if step_match:
                                failure_step_id = step_match.group(1)
                                action_text = step_match.group(2).strip()
                                test_step = action_text  # details should not carry step ID prefix
                                # Extract timestamp from the failure step line if present using best candidate selection
                                consider_timestamp(find_best_timestamp(next_line))
                            else:
                                # Look for common failure indicators in other formats
                                if any(keyword in next_line_lower for keyword in
                                       ["condition", "value", "expected", "actual", "mismatch", "not found",
                                        "exception", "error", "failed to", "failed"]):
                                    if not re.match(r'^\d+\.\d+', next_line):
                                        # Extract step number if it starts with numbers
                                        step_num_match = re.match(r'^(\d+(?:\.\d+)*)', next_line.strip())
                                        if step_num_match:
                                            failure_step_id = step_num_match.group(1)
                                            test_step = next_line[:80]

                        # Look for action keywords for passed cases
                        if verdict_type == "Pass":
                            next_line_lower = next_line.lower()

                            if "execute" in next_line_lower:
                                match = re.search(r'execute\s+(\w+)', next_line_lower)
                                if match:
                                    test_step = match.group(1).capitalize()
                            elif "question" in next_line_lower and "text" in next_line_lower:
                                test_step = "Question/Text"
                            elif "await" in next_line_lower or "wait" in next_line_lower:
                                test_step = "Await Value Match"
                            elif "resume" in next_line_lower:
                                test_step = "Resume"
                            elif "set" in next_line_lower:
                                test_step = "Set"
                            elif "tester" in next_line_lower and "confirmed" in next_line_lower:
                                test_step = "Tester Confirmation"

                    # Add test case if we found a timestamp
                    if timestamp:
                        results[current_fixture]["test_cases"].append({
                            "timestamp": timestamp,
                            "verdict": verdict_type,
                            "details": test_step
                        })

        # Calculate totals
        for fixture_name in results:
            # Use the maximum of: parsed test cases OR initial summary count
            # This preserves the fixture summary count if detailed parsing didn't find individual cases
            parsed_count = len(results[fixture_name]["test_cases"])
            initial_count = results[fixture_name].get("total", 0)
            results[fixture_name]["total"] = max(parsed_count, initial_count)

        return results


    if not st.session_state.selected_files:
        st.info("Select files from the sidebar to show dashboard.")
    elif not dashboard_files:
        st.info("No dashboard-friendly files selected. Choose HTML/HTM/XLSX in sidebar for dashboard details.")
    else:
        st.info(
            "Files selected in the sidebar are available here. Choose only the dashboard file you want to inspect in this tab.")
        dashboard_options = ["--Select File--"] + dashboard_files
        if st.session_state.file_dropdown not in dashboard_options:
            st.session_state.file_dropdown = "--Select File--"
        file_dropdown = st.selectbox("Select a dashboard file", dashboard_options, key="file_dropdown")

        if file_dropdown != "--Select File--":
            with st.spinner("Loading dashboard file..."):
                ensure_file_processed(file_dropdown)
            file_entry = get_uploaded_file_entry(file_dropdown)
            file_bytes = file_entry["bytes"] if file_entry else b""
            chart_type = st.radio("Chart type", ["Pie Chart", "Bar Chart"], index=0)
            bar_orientation = "Vertical"
            if chart_type == "Bar Chart":
                bar_orientation = st.radio("Bar orientation", ["Vertical", "Horizontal"], index=0)
            horizontal_bar = (bar_orientation == "Horizontal")

            if file_dropdown.lower().endswith(".xlsx"):
                data = st.session_state.excel_data_by_file.get(file_dropdown, [])

                if data:
                    columns = ["--Select Column--"] + list(data[0].keys())
                    col = st.selectbox("Column to analyze", columns, index=0)

                    if col != "--Select Column--":
                        counts = get_column_counts(data, col)
                        if counts:
                            fig = plot_pie_chart(counts,
                                                 f"{col} Distribution") if chart_type == "Pie Chart" else plot_bar_chart(
                                counts, f"{col} Distribution", horizontal=horizontal_bar
                            )
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.warning("No data found for the selected column.")
                else:
                    st.warning("No Excel data available for analysis.")

            elif file_dropdown.lower().endswith((".html", ".htm")):
                st.markdown("### 🔐 Login Info")
                login = extract_login_name_from_html(file_bytes)
                st.write("Login Name:", login)

                st.markdown("### 📊 Statistics")
                stats = extract_statistics_from_html(file_bytes)

                if any(v > 0 for v in stats.values()):
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Executed", stats["Executed"])
                    c2.metric("Passed", stats["Passed"])
                    c3.metric("Failed", stats["Failed"])

                    fig = plot_pie_chart(stats, "Statistics") if chart_type == "Pie Chart" else plot_bar_chart(
                        stats, "Statistics", horizontal=horizontal_bar
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("⚠️ Could not extract statistics")

                st.markdown("### 📋 Test Results")
                grouped_results = extract_test_results_grouped_from_html(file_bytes)

                if grouped_results:
                    st.markdown("#### 👉 Executed Test Cases Summary")
                    all_fixtures = sorted(grouped_results.keys())
                    max_cols = 5
                    num_cols = min(len(all_fixtures), max_cols)

                    if num_cols > 0:
                        cols = st.columns(num_cols)
                        for idx, fixture_name in enumerate(all_fixtures[:max_cols]):
                            if fixture_name in grouped_results:
                                total_cases = grouped_results[fixture_name].get("total", 0)
                                pass_count = grouped_results[fixture_name].get("pass", 0)
                                with cols[idx % num_cols]:
                                    st.metric(
                                        label=fixture_name,
                                        value=total_cases,
                                        delta=f"{pass_count} Passed" if total_cases > 0 else None,
                                        delta_color="inverse"
                                    )

                        if len(all_fixtures) > max_cols:
                            st.markdown(f"**And {len(all_fixtures) - max_cols} more fixtures...**")

                    st.markdown("---")
                    st.markdown("#### Full Test Fixtures Overview")

                    fixture_data = []
                    for fixture_name in sorted(grouped_results.keys()):
                        data = grouped_results[fixture_name]
                        fixture_data.append({
                            "Test Fixture": fixture_name,
                            "✅ Pass": data.get("pass", 0),
                            "❌ Fail": data.get("fail", 0),
                            "⚠️ Error": data.get("error", 0),
                            "⏭️ Not Executed": data.get("not executed", 0),
                            "❓ Inconclusive": data.get("inconclusive", 0),
                            "📊 Total": data.get("total", 0)
                        })

                    df = pd.DataFrame(fixture_data)


                    def style_fixture_table(row):
                        styles = [''] * len(row)
                        if '✅ Pass' in df.columns:
                            pass_col_idx = df.columns.get_loc('✅ Pass')
                            if row['✅ Pass'] > 0:
                                styles[pass_col_idx] = 'background-color: #90EE90; color: black; font-weight: bold;'
                        if '❌ Fail' in df.columns:
                            fail_col_idx = df.columns.get_loc('❌ Fail')
                            if row['❌ Fail'] > 0:
                                styles[fail_col_idx] = 'background-color: #FF6B6B; color: white; font-weight: bold;'
                        if '⚠️ Error' in df.columns:
                            error_col_idx = df.columns.get_loc('⚠️ Error')
                            if row['⚠️ Error'] > 0:
                                styles[error_col_idx] = 'background-color: #FF8C8C; color: white; font-weight: bold;'
                        if '⏭️ Not Executed' in df.columns:
                            notexec_col_idx = df.columns.get_loc('⏭️ Not Executed')
                            if row['⏭️ Not Executed'] > 0:
                                styles[notexec_col_idx] = 'background-color: #FFD700; color: black;'
                        if '❓ Inconclusive' in df.columns:
                            inconc_col_idx = df.columns.get_loc('❓ Inconclusive')
                            if row['❓ Inconclusive'] > 0:
                                styles[inconc_col_idx] = 'background-color: #FFA500; color: black;'
                        return styles


                    styled_fixture_df = df.style.apply(style_fixture_table, axis=1)
                    st.dataframe(styled_fixture_df, use_container_width=True)

                    st.markdown("#### View Test Cases by Fixture")
                    fixture_options = sorted(grouped_results.keys())
                    selected_fixture = st.selectbox("Select Test Fixture:", fixture_options, key="fixture_select")

                    if selected_fixture:
                        fixture_info = grouped_results[selected_fixture]
                        st.markdown(f"### 📋 Test Fixture: **{selected_fixture}**")

                        col1, col2, col3, col4, col5 = st.columns(5)
                        col1.metric("✅ Pass", fixture_info.get("pass", 0))
                        col2.metric("❌ Fail", fixture_info.get("fail", 0))
                        col3.metric("⚠️ Error", fixture_info.get("error", 0))
                        col4.metric("⏭️ Not Executed", fixture_info.get("not executed", 0))
                        col5.metric("❓ Inconclusive", fixture_info.get("inconclusive", 0))

                        mode = st.radio("Show Test Cases", ["All", "Passed only", "Failed/Error only"], index=2,
                                        key="test_case_mode")

                        if fixture_info["test_cases"]:
                            test_cases_df = pd.DataFrame()
                            if mode == "All":
                                test_cases_to_show = fixture_info["test_cases"]
                                heading = "Test Cases (All)"
                            elif mode == "Passed only":
                                test_cases_to_show = [
                                    tc for tc in fixture_info["test_cases"]
                                    if tc.get("verdict", "").lower() in ["pass", "passed"]
                                ]
                                heading = "Test Cases (Passed Only)"
                            else:
                                test_cases_to_show = [
                                    tc for tc in fixture_info["test_cases"]
                                    if tc.get("verdict", "").lower() in ["failed", "fail", "error"]
                                ]
                                heading = "Test Cases (Failed/Error Only)"

                            st.markdown(f"#### {heading}")
                            if not test_cases_to_show:
                                st.info("No test cases to display for selected filter.")
                            else:
                                test_cases_df = pd.DataFrame(test_cases_to_show)


                            def color_verdict(val):
                                if val == "Pass":
                                    return "background-color: #90EE90; color: black; font-weight: bold;"
                                if val in ["Fail", "Failed", "Error"]:
                                    return "background-color: #FF6B6B; color: white; font-weight: bold;"
                                if val == "Not Executed":
                                    return "background-color: #FFD700; color: black;"
                                if val == "Inconclusive":
                                    return "background-color: #FFA500; color: black;"
                                return ""


                            if not test_cases_df.empty and "verdict" in test_cases_df.columns:
                                styled_df = test_cases_df.style.map(
                                    lambda x: color_verdict(x) if isinstance(x, str) else "",
                                    subset=["verdict"]
                                )
                                st.dataframe(styled_df, use_container_width=True, hide_index=True)
                            elif not test_cases_df.empty:
                                st.dataframe(test_cases_df, use_container_width=True, hide_index=True)

                        verdict_counts = {
                            "Pass": fixture_info.get("pass", 0),
                            "Fail": fixture_info.get("fail", 0),
                            "Error": fixture_info.get("error", 0),
                            "Not Executed": fixture_info.get("not executed", 0),
                            "Inconclusive": fixture_info.get("inconclusive", 0)
                        }
                        verdict_counts = {k: v for k, v in verdict_counts.items() if v > 0}

                        if verdict_counts:
                            fig = plot_pie_chart(verdict_counts,
                                                 f"Verdict Distribution - {selected_fixture}") if chart_type == "Pie Chart" else plot_bar_chart(
                                verdict_counts, f"Verdict Distribution - {selected_fixture}", horizontal=horizontal_bar
                            )
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("No structured test results found.")

# -------------------------------
# TAB 3: COMPARE
# -------------------------------
# Compare tab:
# Lets users pick 2+ selected files, generate inline word-level differences,
# and download the comparison results as an Excel file.
if active_main_tab == "📂 Compare":
    compare_header_col, compare_reset_col = st.columns([8, 1])
    with compare_header_col:
        st.subheader("Compare Files")
    with compare_reset_col:
        if st.button("🔄 Reset", key="reset_compare_selection", use_container_width=True):
            st.session_state.compare_file_selection = []
            st.session_state.compare_result_html = None
            st.session_state.compare_result_excel_bytes = None
            st.session_state.compare_result_files = []
            st.rerun()

    st.info("Select files in the sidebar to make them available here, then choose only the files you want to compare in this tab.")
    show_current_sidebar_selection()
    show_help_popup('compare', st.session_state.selected_files)
    render_file_context_card("Compare File Context", st.session_state.selected_files, st.session_state.compare_file_selection)

    st.markdown("**Comparison options:**")
    st.markdown(
        "- Inline word-level diff (sequence-aware highlighting)\n"
        "- Side-by-side line diff for direct file comparison\n"
        "- Word presence summary across selected files\n"
        "- Downloadable Excel comparison workbook\n"
        "- Supports PDF, DOCX, PPTX, XLSX, HTML, TXT, CAN/CAPL formats"
    )

    # Use selected files in multiselect (user must choose from selected_files independently)
    st.session_state.compare_file_selection = [
        file_name for file_name in st.session_state.compare_file_selection
        if file_name in st.session_state.selected_files
    ]
    selected_files_for_comparison = st.multiselect(
        "Choose files to compare",
        options=st.session_state.selected_files,
        default=st.session_state.compare_file_selection,
        key="compare_file_selection"
    )

    compare_mode = st.selectbox(
        "Comparison mode",
        ["Exact inline word diff", "Side-by-side line diff", "Word presence summary"],
        index=0,
        key="compare_mode"
    )

    reference_file = None
    # Reference file selection is no longer shown; the first selected file is used automatically for comparison baselines.
    if len(selected_files_for_comparison) == 1:
        st.info("Select at least two files to compare.")

    compare_clicked = st.button("Compare Selected Files", key="run_compare_button", use_container_width=True)

    if compare_clicked:
        if len(selected_files_for_comparison) >= 2:
            with st.spinner("Loading files for comparison..."):
                ensure_files_processed(selected_files_for_comparison)

            selected_texts = {}
            for f in selected_files_for_comparison:
                raw_text = st.session_state.file_texts.get(f, "")
                selected_texts[f] = raw_text if isinstance(raw_text, str) else str(raw_text)

            with st.spinner("Loading comparison results..."):
                html_diff = highlight_multi_file_differences(
                    selected_texts,
                    comparison_mode=compare_mode,
                    reference_file=reference_file
                )
                excel_io = generate_word_level_comparison_excel(selected_texts)

            st.session_state.compare_result_html = html_diff
            st.session_state.compare_result_excel_bytes = excel_io.getvalue()
            st.session_state.compare_result_files = selected_files_for_comparison.copy()
        else:
            st.warning("Select at least two files to compare.")

    if st.session_state.compare_result_html and st.session_state.compare_result_files:
        st.info("Compared files: " + ", ".join(st.session_state.compare_result_files))
        st.markdown(f"### Comparison Results ({len(st.session_state.compare_result_files)} files)")
        st.components.v1.html(st.session_state.compare_result_html, height=800, scrolling=True)

        st.markdown("### Download Excel Comparison")
        st.download_button(
            "Download Comparison Excel",
            st.session_state.compare_result_excel_bytes,
            file_name="comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    elif len(selected_files_for_comparison) < 2:
        st.info("Select at least two files to compare.")

# -------------------------------
# TAB 4: CAPL
# -------------------------------
# CAPL tab:
# Dedicated to CAPL file selection, live editing, compile/analyze checks, issue
# reporting, and optional AI-assisted fix generation for CAPL scripts.
if active_main_tab == "🧠 CAPL":
    capl_header_col, capl_reset_col = st.columns([8, 1])
    with capl_header_col:
        st.subheader("⚙️ CAPL Compiler & Analyzer")
    with capl_reset_col:
        if st.button("🔄 Reset", key="reset_capl_selection", use_container_width=True):
            st.session_state.selected_capl_file = "--Select CAPL file--"
            st.session_state.capl_last_analyzed_file = None
            st.session_state.capl_last_issues = None
            st.rerun()

    st.info(
        "Use sidebar selection to make CAPL source files available here. Then choose the CAPL file you want only in this tab.")
    show_current_sidebar_selection()
    show_help_popup('capl', [f for f in st.session_state.selected_files if f.lower().endswith((".can", ".txt"))])

    # Filter selected files for CAPL analysis
    capl_selectable_files = [f for f in st.session_state.selected_files if f.lower().endswith((".can", ".txt"))]
    active_capl_files = [] if st.session_state.selected_capl_file == "--Select CAPL file--" else [st.session_state.selected_capl_file]
    render_file_context_card("CAPL File Context", capl_selectable_files, active_capl_files)

    with st.expander("✍️ Create New CAPL Script", expanded=False):
        st.text_input("CAPL file name", key="capl_editor_name", help="Enter a file name ending with .can")
        st.text_area("Write CAPL code", key="capl_editor_code", height=260)
        live_editor_issues = analyze_capl_code_with_suggestions(st.session_state.capl_editor_code)

        st.markdown("### Live CAPL Preview")
        st.markdown(
            render_capl_code_with_highlights(st.session_state.capl_editor_code, live_editor_issues),
            unsafe_allow_html=True
        )

        st.markdown("### Live Issues & Suggestions")
        render_capl_issue_table(live_editor_issues)

        editor_ai_cols = st.columns([1, 4])
        with editor_ai_cols[0]:
            editor_ai_fix_clicked = st.button("🤖 Suggest Fix", key="capl_editor_ai_fix_button")

        if editor_ai_fix_clicked:
            llm = load_llm()
            editor_prompt = f"""
            You are a CAPL expert. Here is some CAPL code with errors. Please provide the corrected version of the code that fixes all syntax and logical errors. Only output the corrected CAPL code, nothing else.

            Code:
            {st.session_state.capl_editor_code}
            """
            if llm is None:
                st.error("AI fix feature is unavailable because model backend could not be initialized.")
            else:
                with st.spinner("Generating CAPL fix suggestion..."):
                    try:
                        st.session_state.capl_editor_ai_fix = llm.invoke(editor_prompt).strip()
                    except Exception as exc:
                        st.error(f"AI suggestion failed: {exc}")
                        st.session_state.capl_editor_ai_fix = ""

        if st.session_state.capl_editor_ai_fix:
            st.markdown("### Suggested Corrected CAPL Code")
            fixed_editor_issues = analyze_capl_code_with_suggestions(st.session_state.capl_editor_ai_fix)
            st.markdown(
                render_capl_code_with_highlights(st.session_state.capl_editor_ai_fix, fixed_editor_issues),
                unsafe_allow_html=True
            )
            if st.button("Use Suggested Fix", key="capl_editor_use_ai_fix"):
                st.session_state.capl_editor_code = st.session_state.capl_editor_ai_fix
                st.rerun()

        if st.button("💾 Save New CAPL Script"):
            new_file_name = st.session_state.capl_editor_name.strip()
            if not new_file_name:
                st.warning("Enter a file name for the CAPL script.")
            else:
                if not new_file_name.lower().endswith(".can"):
                    new_file_name += ".can"

                file_bytes = st.session_state.capl_editor_code.encode("utf-8")
                existing_index = next(
                    (idx for idx, file_info in enumerate(st.session_state.uploaded_files) if
                     file_info["name"] == new_file_name),
                    None
                )

                if existing_index is None:
                    st.session_state.uploaded_files.append({"name": new_file_name, "bytes": file_bytes})
                else:
                    st.session_state.uploaded_files[existing_index] = {"name": new_file_name, "bytes": file_bytes}

                st.session_state.file_texts[new_file_name] = st.session_state.capl_editor_code
                if new_file_name not in st.session_state.selected_files:
                    st.session_state.selected_files.append(new_file_name)
                st.session_state.capl_last_analyzed_file = None
                st.session_state.capl_last_issues = None
                st.session_state.capl_editor_ai_fix = ""
                st.success(f"Saved {new_file_name} and added it to the selected files.")

    use_all_txt = st.checkbox("Include all .txt files as CAPL")
    with st.spinner("Loading CAPL files..."):
        ensure_files_processed(capl_selectable_files)

    if use_all_txt:
        capl_files = [
            f for f in capl_selectable_files
            if f.lower().endswith((".can", ".txt"))
        ]
    else:
        capl_files = [
            f for f in capl_selectable_files
            if f.lower().endswith((".can", ".txt")) and
               is_capl_code(st.session_state.file_texts.get(f, ""))
        ]

    if not capl_files:
        st.warning("Upload/select CAPL (.can/.capl) files")
    else:
        capl_options = ["--Select CAPL file--"] + capl_files
        if st.session_state.selected_capl_file not in capl_options:
            st.session_state.selected_capl_file = "--Select CAPL file--"
        selected_capl = st.selectbox("Select CAPL file", capl_options, key="selected_capl_file")
        capl_selected = selected_capl != "--Select CAPL file--"
        action_cols = st.columns([1, 1, 4])
        with action_cols[0]:
            analyze_clicked = st.button("🚀 Compile & Analyze", disabled=not capl_selected)
        with action_cols[1]:
            clear_analysis = st.button("🧹 Clear Analysis", disabled=not capl_selected)

        if not capl_selected:
            st.info("Choose a CAPL file to view and analyze.")
        else:
            if clear_analysis:
                st.session_state.capl_last_analyzed_file = None
                st.session_state.capl_last_issues = None

            code = st.session_state.file_texts[selected_capl]
            issues = []

            if analyze_clicked:
                issues = analyze_capl_code_with_suggestions(code)
                st.session_state.capl_last_analyzed_file = selected_capl
                st.session_state.capl_last_issues = issues

            if st.session_state.capl_last_analyzed_file == selected_capl and st.session_state.capl_last_issues is not None:
                issues = st.session_state.capl_last_issues

            st.markdown("### 📄 CAPL Code")
            st.markdown(render_capl_code_with_highlights(code, issues), unsafe_allow_html=True)

            if st.session_state.capl_last_analyzed_file == selected_capl and st.session_state.capl_last_issues is not None:
                st.markdown("### 🛠️ Issues Found")
                render_capl_issue_table(issues)

                # 🤖 AI Suggestions
                if st.checkbox("🤖 AI Suggestions"):
                    llm = load_llm()
                    if llm is None:
                        st.error("AI fix feature is unavailable because model backend could not be initialized.")
                    else:
                        prompt = f"""
                        You are a CAPL expert. Here is some CAPL code with errors. Please provide the corrected version of the code that fixes all syntax and logical errors. Only output the corrected CAPL code, nothing else.

                        Code:
                        {code}
                        """
                        with st.spinner("Analyzing and fixing with AI..."):
                            try:
                                response = llm.invoke(prompt)
                                corrected_code = response.strip()
                                # Update the code in session state
                                st.session_state.file_texts[selected_capl] = corrected_code
                                code = corrected_code
                                issues = analyze_capl_code_with_suggestions(code)
                                st.session_state.capl_last_analyzed_file = selected_capl
                                st.session_state.capl_last_issues = issues
                                st.success("✅ Code corrected by AI!")
                                st.markdown("### 📄 Corrected CAPL Code")
                                st.markdown(render_capl_code_with_highlights(code, issues), unsafe_allow_html=True)
                                if issues:
                                    st.warning("⚠️ Some issues remain:")
                                render_capl_issue_table(issues)
                            except Exception as exc:
                                st.error(f"AI suggestion failed: {exc}")
