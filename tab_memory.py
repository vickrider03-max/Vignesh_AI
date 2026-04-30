# ==============================
# PERSISTENT PER-TAB MEMORY
# Namespaced memory store that survives Streamlit reruns and avoids tab overlap.
# ==============================
import copy
import streamlit as st


DEFAULT_TAB_MEMORY = {
    "chat": {
        "messages": [],
        "context": {},
        "history": [],
    },
    "dashboard": {
        "selected_file": None,
        "filters": {},
        "history": [],
    },
    "compare": {
        "selected_files": [],
        "last_result": None,
        "history": [],
    },
    "capl": {
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
