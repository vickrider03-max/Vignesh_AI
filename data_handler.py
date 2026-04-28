import re
from difflib import SequenceMatcher
from io import BytesIO

import pandas as pd
import streamlit as st

import data_logic


def get_file_entry(uploaded_files, file_name):
    try:
        for file_info in uploaded_files:
            if file_info.get("name") == file_name:
                return file_info
        return None
    except Exception as exc:
        st.error(f"Error in data_handler.py at function get_file_entry: {exc}")
        return None


@st.cache_data(show_spinner=False)
def load_file_text(file_name, file_bytes):
    try:
        return data_logic.extract_text(file_name, file_bytes)
    except Exception as exc:
        st.error(f"Error in data_handler.py at function load_file_text: {exc}")
        return ""


@st.cache_data(show_spinner=False)
def load_tabular_dataframe(file_name, file_bytes):
    try:
        file_name_lower = file_name.lower()
        if file_name_lower.endswith(".csv"):
            return pd.read_csv(BytesIO(file_bytes))
        if file_name_lower.endswith((".xlsx", ".xls")):
            rows = data_logic.extract_excel_data(file_name, file_bytes)
            return pd.DataFrame(rows)
        return pd.DataFrame()
    except Exception as exc:
        st.error(f"Error in data_handler.py at function load_tabular_dataframe: {exc}")
        return pd.DataFrame()


def process_selected_files(uploaded_files, selected_files):
    try:
        processed = {}
        for file_name in selected_files:
            file_info = get_file_entry(uploaded_files, file_name)
            if not file_info:
                continue
            file_text = load_file_text(file_name, file_info["bytes"])
            processed[file_name] = file_text
            st.session_state.setdefault("file_texts", {})[file_name] = file_text
        return processed
    except Exception as exc:
        st.error(f"Error in data_handler.py at function process_selected_files: {exc}")
        return {}


def search_selected_text(processed_files, query):
    try:
        query_pattern = re.compile(re.escape(query), re.IGNORECASE)
        results = {}
        for file_name, text in processed_files.items():
            matches = []
            for line_number, line in enumerate(str(text).splitlines(), 1):
                if query_pattern.search(line):
                    matches.append((line_number, data_logic.normalize_extracted_line(line)))
            if matches:
                results[file_name] = matches
        return results
    except Exception as exc:
        st.error(f"Error in data_handler.py at function search_selected_text: {exc}")
        return {}


def compare_texts(left_text, right_text):
    try:
        similarity = SequenceMatcher(None, str(left_text), str(right_text)).ratio()
        return {"similarity_percent": round(similarity * 100, 2)}
    except Exception as exc:
        st.error(f"Error in data_handler.py at function compare_texts: {exc}")
        return {"similarity_percent": 0}


def _keyword_list(text, limit=8):
    try:
        stopwords = getattr(data_logic, "SUMMARY_STOPWORDS", set()) or {
            "the", "and", "for", "with", "that", "this", "from", "document", "page", "table"
        }
        words = re.findall(r"[A-Za-z][A-Za-z0-9_+\-/]{2,}", str(text))
        counts = {}
        for word in words:
            lowered = word.lower()
            if len(lowered) <= 3 or lowered in stopwords:
                continue
            counts[lowered] = counts.get(lowered, 0) + 1
        return [word.title() for word, _ in sorted(counts.items(), key=lambda item: item[1], reverse=True)[:limit]]
    except Exception as exc:
        st.error(f"Error in data_handler.py at function _keyword_list: {exc}")
        return []


def build_product_documentation_summary(file_name, text):
    try:
        clean_lines = [
            data_logic.normalize_extracted_line(line)
            for line in str(text).splitlines()
            if 20 <= len(line.strip()) <= 220
        ]
        keywords = _keyword_list(text)
        topic = ", ".join(keywords[:4]) if keywords else file_name
        sample_lines = clean_lines[:8]

        def bullets(items):
            return "\n".join(f"- {item}" for item in items if item)

        return "\n".join([
            "#### Overview",
            bullets([
                f"This document is a reference around {topic}.",
                "Its purpose is to help readers understand the subject, its context, and how to use the information practically.",
            ]),
            "#### Core Concept",
            bullets([
                "The document is reorganized by meaning instead of page order.",
                "Related details are grouped into purpose, structure, capabilities, workflow, and usage value.",
            ]),
            "#### Architecture / Structure",
            bullets(sample_lines[:4] or ["The extracted content does not expose a clear structure."]),
            "#### Key Capabilities",
            bullets([
                "Summarizes functional and technical information.",
                "Highlights useful reference details without copying document text.",
                "Supports quick review, search, comparison, and tabular preview workflows.",
            ]),
            "#### Components / Modules",
            bullets([f"{keyword}: major referenced concept or component." for keyword in keywords[:5]]),
            "#### Workflow / How It Is Used",
            bullets([
                "Upload and select the relevant document.",
                "Extract readable text and structured data.",
                "Review the reorganized summary, search key terms, or inspect extracted content.",
            ]),
            "#### Practical Use Cases",
            bullets([
                "Engineering reference and onboarding.",
                "Document review and quick understanding.",
                "Troubleshooting, validation, reporting, and knowledge extraction.",
            ]),
            "#### Key Takeaways",
            bullets(([
                f"Main focus areas: {', '.join(keywords[:5])}.",
                "The highest value is turning scattered details into usable reference knowledge.",
            ] + sample_lines[:3])[:5]),
        ])
    except Exception as exc:
        st.error(f"Error in data_handler.py at function build_product_documentation_summary: {exc}")
        return "Summary could not be generated."
