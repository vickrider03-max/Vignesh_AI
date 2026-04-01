import difflib, html, re, hashlib
from collections import Counter, defaultdict
from datetime import datetime
from io import BytesIO

import docx, openpyxl, pdfplumber, streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill
from pptx import Presentation
from bs4 import BeautifulSoup
from transformers import pipeline
import plotly.express as px
import plotly.graph_objects as go

from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.llms import HuggingFacePipeline
from langchain_community.vectorstores import FAISS
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.runnables import RunnablePassthrough
from langchain_text_splitters import RecursiveCharacterTextSplitter

# -------------------------------
# STREAMLIT PAGE CONFIG
# -------------------------------
st.set_page_config(page_title="Vignesh_AI", layout="wide")
st.title("🤖 Vignesh_AI")

# -------------------------------
# SESSION STATE INITIALIZATION
# -------------------------------
for key in ["uploaded_files", "selected_files", "file_texts", "excel_data_by_file", "vector_stores", "messages", "ask_messages"]:
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

# -------------------------------
# FILE UPLOAD & MANAGEMENT (SIDEBAR)
# -------------------------------
with st.sidebar:
    if st.session_state.is_authenticated:
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

        st.success(f"Logged in as {st.session_state.logged_in_username} ({st.session_state.user_role})")
        if st.button("Logout"):
            st.session_state.is_authenticated = False
            st.session_state.logged_in_username = ""
            st.session_state.user_role = None
            st.rerun()

        if st.session_state.user_role == "creator":
            st.header("Upload Documents")
            new_files = st.file_uploader(
                "Upload PDF, DOCX, TXT, PPTX, XLSX, HTML,CAPL",
                type=["pdf", "docx", "txt", "pptx", "xlsx", "html", "htm", "capl", "can"],
                accept_multiple_files=True,
                key=f"file_uploader_{st.session_state.file_uploader_key}"
            )

            if new_files:
                existing_names = {f["name"] for f in st.session_state.uploaded_files}
                for file in new_files:
                    if file.name not in existing_names:
                        file_bytes = file.read()
                        st.session_state.uploaded_files.append({"name": file.name, "bytes": file_bytes})

            st.markdown("---")
            st.markdown("### Uploaded files")
            for file_dict in st.session_state.uploaded_files[:]:
                cols = st.columns([0.88, 0.12], vertical_alignment="center")
                with cols[0]:
                    checked = file_dict["name"] in st.session_state.selected_files
                    new_checked = st.checkbox(
                        file_dict["name"],
                        value=checked,
                        key=f"select_{file_dict['name']}"
                    )
                if new_checked and file_dict["name"] not in st.session_state.selected_files:
                    st.session_state.selected_files.append(file_dict["name"])
                elif not new_checked and file_dict["name"] in st.session_state.selected_files:
                    st.session_state.selected_files.remove(file_dict["name"])
                with cols[1]:
                    if st.button("X", key=f"del_{file_dict['name']}", help=f"Delete {file_dict['name']}", type="tertiary"):
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

            st.markdown("---")
            if st.button("Clear All Files"):
                for key in ["uploaded_files", "selected_files", "file_texts", "excel_data_by_file", "vector_stores",
                            "messages"]:
                    st.session_state[key].clear()
                st.session_state.capl_last_analyzed_file = None
                st.session_state.capl_last_issues = None
                st.session_state.file_uploader_key += 1
                st.rerun()

            st.markdown("---")
            st.subheader("Creator Login History")
            if st.session_state.login_history:
                st.table(pd.DataFrame(st.session_state.login_history))
        else:
            st.info("As a regular user, this app only exposes the chatbot. Creator-only management is hidden.")
            st.markdown("*Cannot upload files, edit CAPL, or change script behavior.*")


# -------------------------------
# TEXT EXTRACTION
# -------------------------------
@st.cache_data(show_spinner=False)
def extract_text(file_name, file_bytes):
    text = ""
    bio = BytesIO(file_bytes)
    try:
        if file_name.endswith(".pdf"):
            with pdfplumber.open(bio) as pdf:
                text = "\n".join(p.extract_text() or "" for p in pdf.pages)
        elif file_name.endswith(".txt"):
            text = bio.read().decode("utf-8", errors="ignore")
        elif file_name.endswith(".can"):
            text = bio.read().decode("utf-8", errors="ignore")
        elif file_name.endswith(".docx"):
            doc = docx.Document(bio)
            text = "\n".join([p.text for p in doc.paragraphs])
        elif file_name.endswith(".pptx"):
            prs = Presentation(bio)
            text = "\n".join(shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text"))
        elif file_name.endswith((".html", ".htm")):
            soup = BeautifulSoup(bio.read(), "html.parser")
            text = soup.get_text(separator="\n")
        elif file_name.endswith(".xlsx"):
            wb = openpyxl.load_workbook(bio, data_only=True)
            text = "\n".join(" ".join(str(c) for c in row if c) for sh in wb for row in sh.iter_rows(values_only=True))
    except:
        text = ""
    return text


@st.cache_data(show_spinner=False)
def extract_excel_data(file_name, file_bytes):
    data = []
    bio = BytesIO(file_bytes)
    try:
        if file_name.endswith(".xlsx"):
            wb = openpyxl.load_workbook(bio, data_only=True)
            for sheet in wb:
                headers = None
                for i, row in enumerate(sheet.iter_rows(values_only=True)):
                    if i == 0:
                        headers = list(row)
                    else:
                        if row and any(cell is not None for cell in row):
                            data.append(dict(zip(headers, row)))
    except:
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
    pipe = pipeline("text2text-generation", model="google/flan-t5-small", max_new_tokens=128, return_full_text=False)
    return HuggingFacePipeline(pipeline=pipe)


def get_uploaded_file_entry(file_name):
    for file_info in st.session_state.uploaded_files:
        if file_info["name"] == file_name:
            return file_info
    return None


def ensure_file_processed(file_name):
    file_info = get_uploaded_file_entry(file_name)
    if not file_info:
        return

    if file_name not in st.session_state.file_texts:
        st.session_state.file_texts[file_name] = extract_text(file_name, file_info["bytes"])

    if file_name.endswith(".xlsx") and file_name not in st.session_state.excel_data_by_file:
        st.session_state.excel_data_by_file[file_name] = extract_excel_data(file_name, file_info["bytes"])


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

CREATOR_USERNAME = "creator"
CREATOR_PASSWORD = "creatorpass"

if not st.session_state.is_authenticated:
    st.subheader("Login")
    login_username = st.text_input("Username")
    login_password = st.text_input("Password", type="password")

    if st.button("Access App"):
        cleaned_username = (login_username or "").strip()
        cleaned_password = (login_password or "").strip()

        if cleaned_username == CREATOR_USERNAME and cleaned_password == CREATOR_PASSWORD:
            st.session_state.is_authenticated = True
            st.session_state.logged_in_username = cleaned_username
            st.session_state.user_role = "creator"
            st.session_state.login_history.append({
                "username": cleaned_username,
                "role": "creator",
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            st.success("Creator access granted.")
            st.rerun()

        elif len(cleaned_username) > 3 and len(cleaned_password) > 3:
            st.session_state.is_authenticated = True
            st.session_state.logged_in_username = cleaned_username
            st.session_state.user_role = "user"
            st.session_state.login_history.append({
                "username": cleaned_username,
                "role": "user",
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            st.success("User access granted.")
            st.rerun()

        else:
            st.error("Username and password must each be more than 3 characters.")

    st.info("Creator should use creator/creatorpass; others use any login. Creator sees admin features.")
    st.stop()

process_selected_files()


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
def highlight_multi_file_differences(file_texts):
    if len(file_texts) < 2:
        return "Select at least two files to compare."

    files = list(file_texts.keys())
    css = """
    <style>
        body { font-family: Arial; margin: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid black; padding: 4px; vertical-align: top; white-space: pre-wrap; }
        th { background-color: #f0f0f0; }
        td.line-number { background-color: #f0f0f0; font-weight: bold; text-align: center; }
        .missing { background-color: #ffcccc; }
        .extra { background-color: #ccffcc; }
        .scrollable { overflow:auto; max-height:800px; }
        p.legend span { display:inline-block; width:20px; height:20px; margin-right:5px; vertical-align:middle; }
    </style>
    """
    html_out = "<html><head>" + css + "</head><body><div class='scrollable'>"
    html_out += "<p class='legend'><b>Legend:</b> <span class='missing'></span> Missing, <span class='extra'></span> Extra</p>"
    html_out += "<table><tr><th>Line #</th>" + "".join(f"<th>{fname}</th>" for fname in files) + "</tr>"

    file_lines = {fname: [html.escape(l) for l in t.splitlines()] for fname, t in file_texts.items()}
    max_lines = max(len(l) for l in file_lines.values())

    for i in range(max_lines):
        html_out += f"<tr><td class='line-number'>{i + 1}</td>"
        line_words_all = {f: set(file_lines[f][i].split()) if i < len(file_lines[f]) else set() for f in files}
        all_words = set().union(*line_words_all.values())
        for f in files:
            words = line_words_all[f]
            highlighted = []
            for w in all_words:
                present_in = [ff for ff, ws in line_words_all.items() if w in ws]
                if w in words and len(present_in) < len(files):
                    highlighted.append(f"<span class='extra'>{w}</span>")
                elif w not in words:
                    highlighted.append(f"<span class='missing'>{w}</span>")
                else:
                    highlighted.append(f"<span>{w}</span>")
            html_out += f"<td>{' '.join(highlighted)}</td>"
        html_out += "</tr>"
    html_out += "</table></div></body></html>"
    return html_out


# -------------------------------
# COMPARE EXCEL HIGHLIGHT
# -------------------------------
def generate_word_level_comparison_excel(file_texts):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comparison"

    files = list(file_texts.keys())
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

            # Highlight differences
            all_words_set = set()
            for f in files:
                if i < len(file_lines[f]):
                    all_words_set.update(file_lines[f][i])
            for col_idx, f in enumerate(files, start=2):
                cell = ws.cell(row=ws.max_row, column=col_idx)
                line_words = file_lines[f][i] if i < len(file_lines[f]) else []
                if w_idx >= len(line_words):
                    cell.fill = red_fill
                elif line_words[w_idx] not in all_words_set - set(line_words):
                    continue
                else:
                    cell.fill = green_fill

    excel_io = BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

# -------------------------------
# CAPL Complier
# -------------------------------

def analyze_capl_code_with_suggestions(code):
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
        if re.match(r'^\s*(if|else if|else|for|while|switch)\b', stripped) and not stripped.endswith('{') and not re.search(r'\)\s*\{', stripped):
            # Check if next line starts with '{'
            if i < len(lines) and not lines[i].strip().startswith('{'):
                issues.append({
                    "line": i,
                    "error": "Missing opening brace after control statement",
                    "suggestion": "Add '{' after the condition or on the next line"
                })

        # Detect missing semicolon
        if not stripped.endswith(";") and not stripped.endswith("{") and not stripped.endswith("}"):
            if not re.match(r'^(on|variables|includes|enum|mstimer|timer|if|else|switch|case|for|while|return)\b', stripped):
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


def render_capl_code_with_highlights(code, issues=None):
    """Render CAPL code with IDE-like line highlighting for detected issues."""
    issues = issues or []
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


# -------------------------------
# TABS
# -------------------------------
if st.session_state.user_role == "creator":
    tab1, tab2, tab3, tab4 = st.tabs(["💬 Chat", "📊 Dashboard", "📂 Compare", "🧠 CAPL"])
else:
    tab1, = st.tabs(["💬 Chat"])

# -------------------------------
# TAB 1: CHAT
# -------------------------------
with tab1:
    st.subheader("Chat with Selected Documents")
    if st.session_state.selected_files:
        combined_text = "\n".join([st.session_state.file_texts[f] for f in st.session_state.selected_files])
        combined_vs = get_combined_vector_store(st.session_state.selected_files)
        retriever = combined_vs.as_retriever(search_kwargs={"k": 3})
        llm = load_llm()
        prompt = ChatPromptTemplate.from_messages([
            ("system",
             "You are an intelligent document assistant. Answer ONLY using context.\nIf not found, say 'Not available in documents.'\nContext:\n{context}"),
            ("human", "{question}")
        ])
        chain = ({"context": retriever | (lambda x: '\n'.join(x)), "question": RunnablePassthrough()} | prompt | llm)

        user_input = st.chat_input("Ask something... (type 'clear' to reset chat)")
        if user_input:
            if user_input.strip().lower() == "clear":
                st.session_state.messages = []
                st.success("✅ Chat cleared!")
            else:
                st.session_state.messages.append({"role": "user", "content": user_input})
                with st.spinner("Thinking..."):
                    # Word count queries
                    if any(t in user_input.lower() for t in ["how many", "count", "number of", "occurrences"]):
                        match = re.search(r"'(.*?)'|\"(.*?)\"", user_input)
                        if match:
                            word = match.group(1) or match.group(2)
                            count = len(
                                re.findall(rf'(?<![\w-]){re.escape(word)}(?![\w-])', combined_text, re.IGNORECASE))
                            response = f"🔢 The word/phrase '{word}' appears {count} times."
                        else:
                            response = "⚠️ Specify the word/phrase in quotes."
                    elif "analyze" in user_input.lower() or "summary" in user_input.lower():
                        result = ""
                        for f in st.session_state.selected_files:
                            words = re.findall(r'\w+', st.session_state.file_texts[f].lower())
                            most_common = Counter(words).most_common(10)
                            result += f"📄 **{f}**: Total words {len(words)}, Unique {len(set(words))}, Top {most_common}\n\n"
                        response = result
                    elif "compare" in user_input.lower():
                        selected_texts = {f: st.session_state.file_texts[f] for f in st.session_state.selected_files}
                        response = highlight_multi_file_differences(selected_texts)
                    else:
                        response = str(chain.invoke(user_input))
                    st.session_state.messages.append({"role": "assistant", "content": response})

        for msg in st.session_state.messages:
            role = "🧑" if msg["role"] == "user" else "🤖"
            st.markdown(f"{role} {msg['content']}", unsafe_allow_html=True)
    else:
        st.info("Upload and select files to chat.")

# -------------------------------
# TAB 2: DASHBOARD
# -------------------------------
if st.session_state.user_role == "creator":
    with tab2:
        st.subheader("Dashboard")
    selected_files = st.session_state.selected_files
    dashboard_files = [
        f for f in selected_files
        if f.lower().endswith((".html", ".htm", ".xlsx"))
    ]


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
            results[fixture_name]["total"] = len(results[fixture_name]["test_cases"])

        return results


    if not selected_files:
        st.info("Select files to show dashboard.")
    elif not dashboard_files:
        st.info("Select an HTML or Excel file to view dashboard details.")
    else:
        file_dropdown = st.selectbox("Select a dashboard file", ["--Select File--"] + dashboard_files)

        if file_dropdown != "--Select File--":
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
                            fig = plot_pie_chart(counts, f"{col} Distribution") if chart_type == "Pie Chart" else plot_bar_chart(
                                counts, f"{col} Distribution"
                            )
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.warning("No data found for the selected column.")
                else:
                    st.warning("No Excel data available for analysis.")

            elif file_dropdown.lower().endswith((".html", ".htm")):
                soup = BeautifulSoup(BytesIO(file_bytes), "html.parser")

                st.markdown("### 🔐 Login Info")
                login = extract_login_name_from_html(file_bytes)
                st.write("Login Name:", login)

                with st.expander("🔍 Debug (Raw Text Preview)"):
                    st.text(file_bytes[:2000].decode('utf-8', errors='ignore'))

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

                st.markdown("### 🧪 Test Results")
                grouped_results = extract_test_results_grouped(soup)

                with st.expander("🔍 DEBUG - Parsed Results"):
                    debug_info = st.session_state.get('debug_parse_info', {})
                    st.write(f"**Total lines extracted:** {debug_info.get('total_lines', 0)}")
                    st.write(f"**Test Fixtures found:** {len(grouped_results)}")
                    st.write(f"**Total test cases:** {sum(r.get('total', 0) for r in grouped_results.values())}")

                    if debug_info.get('fixtures_found'):
                        st.write("### Fixtures Detected:")
                        for fixture_info in debug_info['fixtures_found']:
                            st.write(f"- **{fixture_info['name']}** (at line {fixture_info['line_index']})")

                    if debug_info.get('sample_lines'):
                        st.write("### Sample Lines After Fixture Headers:")
                        for sample in debug_info['sample_lines']:
                            st.write(f"**After {sample['fixture']}:**")
                            for l in sample['lines'][:15]:
                                st.code(l, language="text")

                    st.write("### Parsed Test Results:")
                    st.json(grouped_results)

                if grouped_results:
                    st.markdown("#### 🧪 Executed Test Cases Summary")
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

                        mode = st.radio("Show Test Cases", ["All", "Passed only", "Failed/Error only"], index=1,
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

                                styled_df = test_cases_df.style.applymap(
                                    lambda x: color_verdict(x) if isinstance(x, str) else "",
                                    subset=["verdict"]
                                )
                                st.dataframe(styled_df, use_container_width=True, hide_index=True)

                        verdict_counts = {
                            "Pass": fixture_info.get("pass", 0),
                            "Fail": fixture_info.get("fail", 0),
                            "Error": fixture_info.get("error", 0),
                            "Not Executed": fixture_info.get("not executed", 0),
                            "Inconclusive": fixture_info.get("inconclusive", 0)
                        }
                        verdict_counts = {k: v for k, v in verdict_counts.items() if v > 0}

                        if verdict_counts:
                            fig = plot_pie_chart(verdict_counts, f"Verdict Distribution - {selected_fixture}") if chart_type == "Pie Chart" else plot_bar_chart(
                                verdict_counts, f"Verdict Distribution - {selected_fixture}", horizontal=True
                            )
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("No structured test results found.")

# -------------------------------
# TAB 3: COMPARE
# -------------------------------
if st.session_state.user_role == "creator":
    with tab3:
        st.subheader("Compare Files")

    # Use only filenames in multiselect (uploaded_files stores dicts with name/bytes)
    uploaded_file_names = [f.get("name") for f in st.session_state.uploaded_files if isinstance(f, dict) and "name" in f]

    # No default selection, user must actively choose files
    selected_files_for_comparison = st.multiselect(
        "Select files to compare",
        options=uploaded_file_names,
        default=[],
        key="selected_files_for_comparison"
    )

    if len(selected_files_for_comparison) >= 2:
        ensure_files_processed(selected_files_for_comparison)

        selected_texts = {}
        for f in selected_files_for_comparison:
            raw_text = st.session_state.file_texts.get(f, "")
            selected_texts[f] = raw_text if isinstance(raw_text, str) else str(raw_text)

        st.markdown("### Inline Word-Level Comparison")
        html_diff = highlight_multi_file_differences(selected_texts)
        st.components.v1.html(html_diff, height=800, scrolling=True)

        st.markdown("### Download Excel Comparison")
        excel_io = generate_word_level_comparison_excel(selected_texts)
        st.download_button("Download Comparison Excel", excel_io, file_name="comparison.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    else:
        st.info("Select at least two files to compare.")

# -------------------------------
# TAB 4: CAPL
# -------------------------------

if st.session_state.user_role == "creator":
    with tab4:
        st.subheader("⚙️ CAPL Compiler & Analyzer")
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
            with st.spinner("Generating CAPL fix suggestion..."):
                st.session_state.capl_editor_ai_fix = llm.invoke(editor_prompt).strip()

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
                    (idx for idx, file_info in enumerate(st.session_state.uploaded_files) if file_info["name"] == new_file_name),
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

    if use_all_txt:
        capl_files = [
            f for f in st.session_state.selected_files
            if f.lower().endswith((".can", ".txt"))
        ]
    else:
        capl_files = [
            f for f in st.session_state.selected_files
            if f.lower().endswith((".can", ".txt")) and
               is_capl_code(st.session_state.file_texts.get(f, ""))
        ]

    if not capl_files:
        st.warning("Upload/select CAPL (.can/.capl) files")
    else:
        ensure_files_processed(capl_files)
        capl_options = ["--Select CAPL file--"] + capl_files
        selected_capl = st.selectbox("Select CAPL file", capl_options, index=0)
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
                    prompt = f"""
                    You are a CAPL expert. Here is some CAPL code with errors. Please provide the corrected version of the code that fixes all syntax and logical errors. Only output the corrected CAPL code, nothing else.

                    Code:
                    {code}
                    """
                    with st.spinner("Analyzing and fixing with AI..."):
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