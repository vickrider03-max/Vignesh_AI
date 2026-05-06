# ==============================
# PERSISTENT PER-TAB MEMORY
# Namespaced memory store that survives Streamlit reruns and avoids tab overlap.
# ==============================
import copy
import streamlit as st


DEFAULT_TAB_MEMORY = {
    "chat": {
        "uploaded_files": [],
        "messages": [],
        "context": {},
        "history": [],
    },
    "dashboard": {
        "uploaded_files": [],
        "selected_file": None,
        "filters": {},
        "history": [],
    },
    "compare": {
        "uploaded_files": [],
        "selected_files": [],
        "last_result": None,
        "history": [],
    },
    "capl": {
        "uploaded_files": [],
        "selected_file": None,
        "issues": [],
        "analysis_cache": {},
    },
}


def init_tab_memory():
    """Initialize isolated memory without overwriting existing tab data."""
    if "tab_memory" not in st.session_state or not isinstance(st.session_state.tab_memory, dict):
        st.session_state.tab_memory = {}
    for tab_name, defaults in DEFAULT_TAB_MEMORY.items():
        if tab_name not in st.session_state.tab_memory or not isinstance(st.session_state.tab_memory[tab_name], dict):
            st.session_state.tab_memory[tab_name] = copy.deepcopy(defaults)
        else:
            for key, default_value in defaults.items():
                st.session_state.tab_memory[tab_name].setdefault(key, copy.deepcopy(default_value))


def get_tab_memory(tab_name):
    init_tab_memory()
    return st.session_state.tab_memory[tab_name]


def tab_memory_get(tab_name, key, default=None):
    return get_tab_memory(tab_name).get(key, default)


def tab_memory_set(tab_name, key, value):
    get_tab_memory(tab_name)[key] = value
    return value


def append_tab_history(tab_name, event):
    memory = get_tab_memory(tab_name)
    memory.setdefault("history", []).append(event)
    memory["history"] = memory["history"][-100:]
    return event


def get_tab_uploaded_files(tab_name):
    """Return the persisted uploaded files for one tab."""
    memory = get_tab_memory(tab_name)
    uploads = memory.setdefault("uploaded_files", [])
    if not isinstance(uploads, list):
        memory["uploaded_files"] = []
    return memory["uploaded_files"]


def remember_tab_upload(tab_name, file_name, file_bytes, status="queued"):
    """Persist uploaded bytes independently from Streamlit's file_uploader widget."""
    uploads = get_tab_uploaded_files(tab_name)
    entry = {
        "name": file_name,
        "bytes": file_bytes,
        "status": status,
    }
    for index, existing in enumerate(uploads):
        if existing.get("name") == file_name:
            uploads[index] = entry
            return entry
    uploads.append(entry)
    return entry


def remove_tab_upload(tab_name, file_name):
    """Remove one persisted file from a tab upload bucket."""
    uploads = get_tab_uploaded_files(tab_name)
    get_tab_memory(tab_name)["uploaded_files"] = [
        upload for upload in uploads
        if upload.get("name") != file_name
    ]


def clear_tab_uploads(tab_name):
    """Clear only one tab's persisted upload bucket."""
    get_tab_memory(tab_name)["uploaded_files"] = []
