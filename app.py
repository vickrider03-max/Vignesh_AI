import html, re, hashlib, os, json, base64, pickle
import uuid
import urllib.parse
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from io import BytesIO
from pytz import timezone

import docx, openpyxl, pdfplumber, streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill
from pptx import Presentation
from bs4 import BeautifulSoup
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

# -------------------------------
# STREAMLIT PAGE CONFIG
# -------------------------------
st.set_page_config(page_title="Vignesh_AI", layout="wide")

# Load preview data from file
load_preview_data()

# Clean up expired preview tokens on app start
cleanup_expired_preview_tokens()

st.title("🤖 Vignesh_AI")
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
        .status-strip {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 10px;
            margin: 8px 0 18px 0;
        }
        .status-tile {
            background: linear-gradient(180deg, #ffffff 0%, #f7fbff 100%);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 12px 14px;
        }
        .status-label {
            color: var(--text-soft);
            font-size: 12px;
            margin-bottom: 4px;
        }
        .status-value {
            color: #173152;
            font-size: 18px;
            font-weight: 600;
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
    </style>
    """,
    unsafe_allow_html=True
)

# -------------------------------
# SESSION STATE INITIALIZATION
# -------------------------------
for key in ["uploaded_files", "selected_files", "file_texts", "excel_data_by_file", "vector_stores", "messages",
            "ask_messages"]:
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
if "compare_file_selection" not in st.session_state:
    st.session_state.compare_file_selection = []
if "file_dropdown" not in st.session_state:
    st.session_state.file_dropdown = "--Select File--"
if "selected_capl_file" not in st.session_state:
    st.session_state.selected_capl_file = "--Select CAPL file--"
if "llm_task" not in st.session_state:
    st.session_state.llm_task = None

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
        
        # Count images
        image_count = 0
        for rel in doc.part.rels.values():
            if "image" in rel.reltype:
                image_count += 1
        if image_count > 0:
            content.append(("IMAGE", f"{image_count} images found in document"))
        
        # Extract paragraphs
        paragraphs = []
        for para in doc.paragraphs:
            if para.text.strip():
                paragraphs.append(para.text)
        
        if paragraphs:
            content.append(("TEXT", "\n\n".join(paragraphs)))
        
        # Extract tables
        for i, table in enumerate(doc.tables):
            table_data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    row_data.append(cell_text)
                table_data.append(row_data)
            
            if table_data:
                table_text = "\n".join([" | ".join(row) for row in table_data])
                content.append(("TABLE", f"Table {i+1}:\n{table_text}"))
    
    except Exception as e:
        content.append(("ERROR", f"DOCX extraction failed: {str(e)}"))
    
    return content


def extract_pptx_content(bio):
    """Extract text, tables, and images from PPTX."""
    content = []
    try:
        prs = Presentation(bio)
        
        content.append(("METADATA", f"Total Slides: {len(prs.slides)}"))
        
        # Count images across all slides
        image_count = 0
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'image'):
                    image_count += 1
        if image_count > 0:
            content.append(("IMAGE", f"{image_count} images found in presentation"))
        
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




def render_document_preview(file_name, file_entry=None):
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

    # Special handling for images
    if file_name_lower.endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp")):
        with st.spinner("Loading preview..."):
            try:
                st.image(file_entry["bytes"], caption=file_name, use_container_width=True)
                # Determine MIME type
                ext = file_name_lower.split('.')[-1]
                mime_type = f"image/{ext}"
                if ext == "jpg":
                    mime_type = "image/jpeg"
                elif ext == "svg":
                    mime_type = "image/svg+xml"
                
                st.download_button(
                    label="📥 Download Image",
                    data=file_entry["bytes"],
                    file_name=file_name,
                    mime=mime_type,
                    key=f"download_image_{file_name}"
                )
            except Exception as e:
                st.error(f"Could not display image: {e}")
                st.download_button(
                    label="📥 Download Image File",
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
                st.dataframe(pd.DataFrame(data).head(20), use_container_width=True, hide_index=True)
            else:
                st.info("No preview data available for this spreadsheet.")
        return

    # For all other files, show comprehensive extracted content
    with st.spinner("Loading preview..."):
        full_content = st.session_state.file_texts.get(file_name, "")
        
        if not full_content.strip():
            st.info("No content could be extracted from this file.")
            return

        # Parse the content into sections
        sections = parse_extracted_content(full_content)
        
        # Display each section
        for section_type, section_title, section_content in sections:
            if section_type == "METADATA":
                with st.expander(f"📋 {section_title}", expanded=False):
                    st.code(section_content, language="text")
            elif section_type == "TEXT":
                st.markdown(f"### {section_title}")
                # Show first 10,000 characters, with option to expand
                if len(section_content) > 10000:
                    st.text_area(
                        f"{section_title} (truncated)",
                        value=section_content[:10000] + "\n\n[... content truncated ...]",
                        height=400,
                        disabled=True,
                        key=f"preview_text_{file_name}_{section_title}"
                    )
                    with st.expander("Show full content"):
                        st.text_area(
                            f"Full {section_title}",
                            value=section_content,
                            height=600,
                            disabled=True,
                            key=f"preview_full_text_{file_name}_{section_title}"
                        )
                else:
                    st.text_area(
                        section_title,
                        value=section_content,
                        height=400,
                        disabled=True,
                        key=f"preview_text_{file_name}_{section_title}"
                    )
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
                                    label="📥 Download as CSV",
                                    data=csv_data,
                                    file_name=f"{section_title.replace(' ', '_')}.csv",
                                    mime="text/csv",
                                    key=f"download_table_{file_name}_{section_title}"
                                )
                            else:
                                st.code(section_content, language="text")
                    except Exception:
                        st.code(section_content, language="text")
            elif section_type == "IMAGE":
                st.info(f"🖼️ {section_content}")
            elif section_type == "ERROR":
                st.error(f"❌ {section_content}")
            elif section_type == "UNSUPPORTED":
                st.warning(f"⚠️ {section_content}")


def parse_extracted_content(content):
    """Parse the extracted content into sections for display."""
    sections = []
    lines = content.split('\n')
    current_section = None
    current_content = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Check for section markers
        if line.startswith('TABLE:'):
            # Save previous section
            if current_section and current_content:
                sections.append(current_section)
            
            # Start new table section
            current_section = ("TABLE", "Table Content", "")
            current_content = []
            
        elif line.startswith('[IMAGE:'):
            # Save previous section
            if current_section and current_content:
                sections.append(current_section)
            
            # Add image info
            sections.append(("IMAGE", "Images", line))
            current_section = None
            current_content = []
            
        elif line.startswith('PDF Metadata:') or line.startswith('Document Metadata:') or line.startswith('Meta Tags:') or line.startswith('Title:') or 'Pages:' in line or 'Slides:' in line or 'sheets:' in line:
            # Save previous section
            if current_section and current_content:
                sections.append(current_section)
            
            # Start metadata section
            if 'Metadata:' in line or 'Tags:' in line:
                section_title = line.split(':')[0] + " Information"
            else:
                section_title = "Document Information"
            
            current_section = ("METADATA", section_title, line)
            current_content = [line]
            
        elif line.startswith('Page ') and 'Text:' in line:
            # Save previous section
            if current_section and current_content:
                sections.append(current_section)
            
            # Start new text section
            current_section = ("TEXT", f"Page {line.split()[1]} Content", "")
            current_content = []
            
        elif line.startswith('Slide ') and ':' in line:
            # Save previous section
            if current_section and current_content:
                sections.append(current_section)
            
            # Start new slide section
            slide_num = line.split(':')[0]
            current_section = ("TEXT", f"{slide_num} Content", "")
            current_content = []
            
        elif line.startswith('Sheet ') and ':' in line:
            # Save previous section
            if current_section and current_content:
                sections.append(current_section)
            
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
    if current_section and current_content:
        if current_section[0] == "TEXT" or current_section[0] == "TABLE":
            content_text = '\n'.join(current_content)
            sections.append((current_section[0], current_section[1], content_text))
        else:
            sections.append(current_section)
    
    # If no sections were created but we have content, create a default text section
    if not sections and content.strip():
        sections.append(("TEXT", "Document Content", content.strip()))
    
    return sections


# Handle browser tab preview links
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

if "preview_token" in query_params and query_params["preview_token"]:
    preview_value = query_params["preview_token"]
    if isinstance(preview_value, list):
        preview_value = preview_value[0] if preview_value else ""
    preview_token = str(preview_value)
    token_data = PREVIEW_TOKENS.get(preview_token)
    preview_file_from_url = token_data['file_name'] if token_data else None

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
    st.title("📄 File Preview")
    if preview_entry is not None:
        st.markdown(
            "<a href='/' style='display:inline-block;padding:10px 18px;border-radius:8px;background:#1f4f91;color:#ffffff;text-decoration:none;'>← Back to app</a>",
            unsafe_allow_html=True,
        )
        st.markdown(f"### {preview_entry['name']}")
        st.markdown("---")
        render_document_preview(preview_entry['name'], file_entry=preview_entry)
    else:
        st.error("Preview file not found in the preview store. Please return to the app and click preview again.")
    st.stop()


# -------------------------------
# FILE UPLOAD & MANAGEMENT (SIDEBAR)
# -------------------------------
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
            </style>
            """,
            unsafe_allow_html=True
        )

        status_message = f"Logged in as {st.session_state.logged_in_username} ({st.session_state.user_role})"
        if creator_timestamp:
            status_message += f"\nLogin time: {creator_timestamp}"
        st.success(status_message)
        if st.button("Logout"):
            # Remove from active users
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
            st.rerun()

        st.header("Upload Documents")
        st.info("1) Upload files. 2) Check the files you need. 3) Switch tabs and work with selected files.")
        new_files = st.file_uploader(
            "Upload PDF, DOCX, TXT, PPTX, XLSX, HTML, CAPL, Images",
            type=["pdf", "docx", "txt", "pptx", "xlsx", "html", "htm", "capl", "can", "png", "jpg", "jpeg", "gif", "bmp", "webp"],
            accept_multiple_files=True,
            key=f"file_uploader_{st.session_state.file_uploader_key}"
        )

        if new_files:
            with st.spinner("Loading uploaded files..."):
                existing_names = {f["name"] for f in st.session_state.uploaded_files}
                for file in new_files:
                    if file.name not in existing_names:
                        file_bytes = file.read()
                        st.session_state.uploaded_files.append({"name": file.name, "bytes": file_bytes})

        st.markdown("---")
        st.markdown("### Uploaded files")
        
        for idx, file_dict in enumerate(st.session_state.uploaded_files[:]):
            cols = st.columns([0.50, 0.20, 0.15, 0.15], vertical_alignment="center")
            with cols[0]:
                checked = file_dict["name"] in st.session_state.selected_files
                new_checked = st.checkbox(
                    file_dict["name"],
                    value=checked,
                    key=f"select_file_{idx}"
                )
            if new_checked and file_dict["name"] not in st.session_state.selected_files:
                st.session_state.selected_files.append(file_dict["name"])
            elif not new_checked and file_dict["name"] in st.session_state.selected_files:
                st.session_state.selected_files.remove(file_dict["name"])
            
            with cols[1]:
                token = str(uuid.uuid4())
                PREVIEW_TOKENS[token] = {'file_name': file_dict["name"], 'timestamp': datetime.now()}
                PREVIEW_STORE[token] = file_dict
                save_preview_data()  # Save to file so preview tab can access it
                st.markdown(
                    f"<a href='./?preview_token={token}' target='_blank' style='display:inline-block;padding:6px 10px;border-radius:8px;background:#1f4f91;color:#ffffff;text-decoration:none;font-size:16px;'>👁️</a>",
                    unsafe_allow_html=True,
                )
            
            with cols[2]:
                # Delete button
                if st.button("🗑️", key=f"del_file_{idx}", help=f"Delete {file_dict['name']}", use_container_width=True):
                    deleted_name = file_dict["name"]
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
@st.cache_data(show_spinner=False)
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


@st.cache_data(show_spinner=False)
def build_file_overview(file_name, text):
    lines = [line.strip() for line in str(text).splitlines() if line.strip()]
    words = re.findall(r"\w+", str(text))
    top_terms = Counter(word.lower() for word in words if len(word) > 2).most_common(8)
    preview_lines = lines[:5]

    summary_parts = [
        f"📄 **{file_name}**",
        f"- Characters: {len(str(text))}",
        f"- Lines: {len(str(text).splitlines())}",
        f"- Words: {len(words)}",
        f"- Non-empty lines: {len(lines)}",
    ]

    if top_terms:
        summary_parts.append(
            "- Top keywords: " + ", ".join(f"{word} ({count})" for word, count in top_terms)
        )

    if preview_lines:
        summary_parts.append("- Preview:")
        summary_parts.extend(f"  {line[:180]}" for line in preview_lines)
    else:
        summary_parts.append("- No readable text could be extracted from this file.")

    return "\n".join(summary_parts)


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

if not st.session_state.is_authenticated and "preview_token" not in query_params:
    st.subheader("Login")
    login_username = st.text_input("Username")
    login_password = st.text_input("Password", type="password")
    st.info("Note: For users, leave the password field blank (empty).")
    has_read_readme = st.checkbox("I have read the README and want to continue")
    st.markdown("### Read Me First")
    st.text_area("README", value=README_TEXT, height=360, disabled=True)

    if st.button("Access"):
        if not has_read_readme:
            st.warning("Please read the README and confirm the checkbox before accessing the app.")
            st.stop()

        cleaned_username = (login_username or "").strip()
        cleaned_password = (login_password or "").strip()

        if cleaned_username == CREATOR_USERNAME and cleaned_password == CREATOR_PASSWORD:
            st.session_state.is_authenticated = True
            st.session_state.logged_in_username = cleaned_username
            st.session_state.user_role = "creator"
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
                "For creator: use username 'creator' and password 'creatorpass'. For users: username >3 chars, password empty.")

    # st.info("Creator should use creator/creatorpass; others use any login. Creator sees admin features.")
    st.stop()


# Files will be processed on-demand per tab when selected


# -------------------------------
# HELPER CHART FUNCTIONS
# -------------------------------
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
@st.cache_data(show_spinner=False)
def highlight_multi_file_differences_cached(file_items):
    if len(file_items) < 2:
        return "Select at least two files to compare."

    files = [fname for fname, _ in file_items]
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
        "<p class='legend'><b>Legend:</b> <span class='match'></span> Exact match, <span class='mismatch'></span> Missing or mismatch</p>",
        "<table><tr><th>Line #</th>",
        "".join(f"<th>{html.escape(fname)}</th>" for fname in files),
        "</tr>",
    ]

    file_lines = {fname: text.splitlines() for fname, text in file_items}
    max_lines = max(len(lines) for lines in file_lines.values())

    for i in range(max_lines):
        html_parts.append(f"<tr><td class='line-number'>{i + 1}</td>")

        line_word_lists = {}
        word_presence = defaultdict(int)
        ordered_words = []

        for fname in files:
            raw_line = file_lines[fname][i] if i < len(file_lines[fname]) else ""
            words = raw_line.split()
            line_word_lists[fname] = words
            seen_words = set()
            for word in words:
                if word not in seen_words:
                    word_presence[word] += 1
                    seen_words.add(word)
                if word not in ordered_words:
                    ordered_words.append(word)

        for fname in files:
            words = set(line_word_lists[fname])
            highlighted = []
            for word in ordered_words:
                escaped_word = html.escape(word)
                if word in words and word_presence[word] == len(files):
                    highlighted.append(f"<span class='match'>{escaped_word}</span>")
                else:
                    highlighted.append(f"<span class='mismatch'>{escaped_word}</span>")
            html_parts.append(f"<td>{' '.join(highlighted) if highlighted else '&nbsp;'}</td>")

        html_parts.append("</tr>")

    html_parts.append("</table></div></body></html>")
    return "".join(html_parts)


def highlight_multi_file_differences(file_texts):
    return highlight_multi_file_differences_cached(tuple((fname, str(text)) for fname, text in file_texts.items()))


# -------------------------------
# COMPARE EXCEL HIGHLIGHT
# -------------------------------
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


def show_current_sidebar_selection():
    selected = st.session_state.get("selected_files", [])
    if selected:
        st.info("Sidebar selected files: " + ", ".join(selected))
    else:
        st.info("No sidebar files selected yet. Upload and select files from the sidebar first.")


def render_status_strip():
    if not st.session_state.get("is_authenticated"):
        return

    available_files = st.session_state.get("selected_files", [])
    role = st.session_state.get("user_role") or "-"
    username = st.session_state.get("logged_in_username") or "-"
    login_entries = st.session_state.get("login_history", [])
    last_login = login_entries[-1]["timestamp"] if login_entries else "-"
    last_login_display = last_login if role == "creator" else "Active Session"

    st.markdown(
        f"""
        <div class="status-strip">
            <div class="status-tile">
                <div class="status-label">User</div>
                <div class="status-value">{html.escape(username)}</div>
            </div>
            <div class="status-tile">
                <div class="status-label">Role</div>
                <div class="status-value">{html.escape(str(role).title())}</div>
            </div>
            <div class="status-tile">
                <div class="status-label">Available Files</div>
                <div class="status-value">{len(available_files)}</div>
            </div>
            <div class="status-tile">
                <div class="status-label">{ "Last Login" if role == "creator" else "Session Status" }</div>
                <div class="status-value" style="font-size:14px;">{html.escape(str(last_login_display))}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


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


def show_help_popup(tab_name, selected_files):
    state_key = ensure_help_popup_state(tab_name)
    st.checkbox("Show query helper popup", key=state_key)

    if not st.session_state[state_key]:
        return

    if not selected_files:
        st.markdown(
            """
            <div style='position:fixed; bottom:14px; right:14px; width:340px; padding:14px; background:#ffffff; border:1px solid #1f4f91; border-radius:12px; box-shadow:0 8px 24px rgba(0,0,0,0.24); z-index:9999;'>
                <h4 style='margin:0 0 6px 0; font-size:15px; color:#1f4f91;'>📘 Document Chat Syntax</h4>
                <p style='margin:0; font-size:13px; color:#253659;'>Select a document first to see targeted query guidance.</p>
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
        extra = "Select 2+ files; use compare shorthand like compare('Error')(doc1, doc2) in natural language."
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
            <h4 style='margin:0 0 6px 0; font-size:15px; color:#1f4f91;'>📘 {tab_name.capitalize()} Query Help</h4>
            <p style='margin:0 0 8px 0; font-size:13px; color:#253659;'>Support for selected file types: {', '.join(sorted(selected_types))}</p>
            <ul style='margin:0 0 8px 18px; padding:0; color:#253659; font-size:13px;'>
                <li><code style='background:#f1f5f9; padding:1px 4px; border-radius:3px;'>count('Approved')</code> - frequency</li>
                <li><code style='background:#f1f5f9; padding:1px 4px; border-radius:3px;'>find('Error')</code> - locate phrases</li>
                <li><code style='background:#f1f5f9; padding:1px 4px; border-radius:3px;'>summarize()</code> - quick summary</li>
            </ul>
            <p style='margin:0; font-size:12px; color:#3c4f7e;'>{extra}</p>
            <p style='margin:5px 0 0 0; font-size:12px; color:#3c4f7e;'>{type_hint}</p>
        </div>
        """,
        unsafe_allow_html=True
    )


# -------------------------------
# TABS
# -------------------------------
render_status_strip()

# All authenticated users get full panel tabs; content can still be permission-aware.
tab1, tab2, tab3, tab4 = st.tabs(["💬 Chat", "📊 Dashboard", "📂 Compare", "🧠 CAPL"])

# -------------------------------
# TAB 1: CHAT
# -------------------------------
with tab1:
    chat_header_col, chat_reset_col = st.columns([8, 1])
    with chat_header_col:
        st.subheader("Chat with Selected Documents")
    with chat_reset_col:
        if st.button("🔄 Reset", key="reset_chat_selection", use_container_width=True):
            st.session_state.chat_file_selection = []
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
                    st.success("✅ Chat cleared!")
                else:
                    st.session_state.messages.append({"role": "user", "content": user_input})
                    with st.spinner("Processing your request..."):
                        # Word count queries
                        if any(t in user_input.lower() for t in ["how many", "count", "number of", "occurrences"]):
                            match = re.search(r"'(.*?)'|\"(.*?)\"", user_input)
                            if match:
                                word = match.group(1) or match.group(2)
                                count = len(
                                    re.findall(rf'(?<![\w-]){re.escape(word)}(?![\w-])', combined_text, re.IGNORECASE))
                                response = f"🔢 The word/phrase '{word}' appears {count} times in the selected documents."
                            else:
                                response = "⚠️ Specify the word/phrase in quotes. Example: count('keyword') or count(\"keyword\")"
                        elif any(term in user_input.lower() for term in ["find", "search", "locate"]) or "highlight" in user_input.lower():
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
                        elif any(term in user_input.lower() for term in ["analyze", "summary", "summarize", "overview"]):
                            result = []
                            llm = load_llm()
                            for f in chat_files:
                                file_text = st.session_state.file_texts.get(f, "")
                                if file_text.strip():
                                    # Try LLM first
                                    if llm:
                                        summary_prompt = f"""Please provide a structured summary of the following document. Extract and organize information into these key sections:

1. **Introduction/Overview**: Main purpose and description of the product/system
2. **Key Features/Advantages**: Main benefits and capabilities
3. **System Requirements**: Hardware/software requirements mentioned
4. **Technical Specifications**: Important technical details, versions, or specifications
5. **Usage Instructions**: How to use or implement the system
6. **Important Notes/Warnings**: Any critical information, limitations, or warnings

Document: {f}
Content: {file_text[:8000]}

Please provide a comprehensive summary organized by these sections. If a section is not mentioned in the document, you can omit it or note that it's not available."""
                                        try:
                                            summary_response = llm.invoke(summary_prompt)
                                            summary_text = str(summary_response).strip()
                                            if summary_text and len(summary_text) > 50 and not summary_text.lower().startswith(f.lower()):
                                                result.append(f"📄 **{f}**\n\n{summary_text}")
                                                continue
                                        except Exception:
                                            pass
                                    
                                    # Fallback: Extract meaningful content manually
                                    lines = [line.strip() for line in file_text.splitlines() if line.strip() and len(line.strip()) > 10]
                                    words = re.findall(r'\w+', file_text)
                                    
                                    # Look for tables
                                    table_lines = [line for line in lines if '|' in line or 'TABLE:' in line]
                                    
                                    summary_parts = [f"📄 **{f}**"]
                                    
                                    if table_lines:
                                        summary_parts.append("\n**Tables Found:**")
                                        summary_parts.extend(f"- {line[:200]}" for line in table_lines[:5])
                                    
                                    # Extract key sentences (lines with important keywords)
                                    key_sentences = []
                                    for line in lines[:20]:  # First 20 lines
                                        if any(keyword in line.lower() for keyword in ['introduction', 'overview', 'summary', 'key', 'important', 'main', 'product', 'information']):
                                            key_sentences.append(line)
                                    
                                    if key_sentences:
                                        summary_parts.append("\n**Key Information:**")
                                        summary_parts.extend(f"- {sent[:200]}" for sent in key_sentences[:5])
                                    
                                    # Basic stats
                                    summary_parts.append(f"\n**Document Statistics:**")
                                    summary_parts.append(f"- Total characters: {len(file_text)}")
                                    summary_parts.append(f"- Estimated words: {len(words)}")
                                    summary_parts.append(f"- Content lines: {len(lines)}")
                                    
                                    # Preview of content
                                    if lines:
                                        summary_parts.append(f"\n**Content Preview:**")
                                        preview = ' '.join(lines[:3])[:500]
                                        summary_parts.append(f"{preview}...")
                                    
                                    result.append('\n'.join(summary_parts))
                                else:
                                    result.append(f"📄 **{f}**\n\nNo readable content found in this document.")
                            
                            response = "\n\n---\n\n".join(result)
                        elif "compare" in user_input.lower():
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
    else:
        st.info("Select files from the sidebar to start chatting.")

# -------------------------------
# TAB 2: DASHBOARD
# -------------------------------
with tab2:
    dashboard_header_col, dashboard_reset_col = st.columns([8, 1])
    with dashboard_header_col:
        st.subheader("Dashboard")
    with dashboard_reset_col:
        if st.button("🔄 Reset", key="reset_dashboard_selection", use_container_width=True):
            st.session_state.file_dropdown = "--Select File--"
            st.rerun()

    show_current_sidebar_selection()

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
            chart_type = st.radio("Chart type", ["Bar Chart", "Pie Chart"])

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
                                counts, f"{col} Distribution"
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
                        stats, "Statistics"
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


                            if "verdict" in test_cases_df.columns:
                                styled_df = test_cases_df.style.map(
                                    lambda x: color_verdict(x) if isinstance(x, str) else "",
                                    subset=["verdict"]
                                )
                                st.dataframe(styled_df, use_container_width=True, hide_index=True)
                            else:
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
                                verdict_counts, f"Verdict Distribution - {selected_fixture}", horizontal=True
                            )
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("No structured test results found.")

# -------------------------------
# TAB 3: COMPARE
# -------------------------------
with tab3:
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
    render_file_context_card("Compare File Context", st.session_state.selected_files, st.session_state.compare_file_selection)

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
                html_diff = highlight_multi_file_differences(selected_texts)
                excel_io = generate_word_level_comparison_excel(selected_texts)

            st.session_state.compare_result_html = html_diff
            st.session_state.compare_result_excel_bytes = excel_io.getvalue()
            st.session_state.compare_result_files = selected_files_for_comparison.copy()
        else:
            st.warning("Select at least two files to compare.")

    if st.session_state.compare_result_html and st.session_state.compare_result_files:
        st.info("Compared files: " + ", ".join(st.session_state.compare_result_files))
        st.markdown(f"### Inline Word-Level Comparison ({len(st.session_state.compare_result_files)} files)")
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

with tab4:
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
