# ==============================
# STATE FIREWALL SYSTEM
# Prevents accidental cross-tab state writes by centralizing write validation.
# This is a compatibility layer: existing legacy code can migrate to firewall_set
# gradually without breaking current Streamlit behavior.
# ==============================
import contextlib
import streamlit as st


TAB_NAMES = ("chat", "dashboard", "compare", "capl")

SHARED_KEYS = {
    "is_authenticated",
    "logged_in_username",
    "user_role",
    "login_history",
    "uploaded_files",
    "selected_files",
    "file_texts",
    "excel_data_by_file",
    "vector_stores",
    "extracted_images",
    "workspace_memory",
    "workspace_memory_loaded",
    "file_uploader_key",
    "active_main_tab",
    "active_tab",
    "tab_colors",
    "context_memory",
    "pending_scroll_anchor",
    "mobile_sidebar_visible",
    "llm_task",
}

TAB_PREFIXES = {
    "chat": ("chat_", "messages", "ask_messages", "input_prefill", "last_streamed_assistant_index"),
    "dashboard": ("dashboard_", "file_dropdown", "fixture_select", "test_case_mode"),
    "compare": ("compare_",),
    "capl": ("capl_", "selected_capl_file", "agent_run_history"),
}


def init_state_firewall():
    """Create firewall metadata without mutating feature state."""
    if "_state_firewall" not in st.session_state:
        st.session_state["_state_firewall"] = {
            "active_scope": "app",
            "violations": [],
            "enforce": False,
        }


def get_active_scope():
    init_state_firewall()
    return st.session_state["_state_firewall"].get("active_scope", "app")


def is_key_allowed(tab_name, key):
    """Return True when a tab may write a session_state key."""
    if tab_name in (None, "", "app"):
        return True
    if key in SHARED_KEYS:
        return True
    prefixes = TAB_PREFIXES.get(tab_name, ())
    return any(str(key).startswith(prefix) for prefix in prefixes)


def record_state_violation(tab_name, key):
    init_state_firewall()
    violation = {"tab": tab_name, "key": str(key)}
    st.session_state["_state_firewall"]["violations"].append(violation)
    st.session_state["_state_firewall"]["violations"] = st.session_state["_state_firewall"]["violations"][-50:]


def firewall_set(tab_name, key, value):
    """Validated write API for migrated code paths."""
    init_state_firewall()
    if not is_key_allowed(tab_name, key):
        record_state_violation(tab_name, key)
        if st.session_state["_state_firewall"].get("enforce"):
            raise PermissionError(f"Tab '{tab_name}' cannot write session key '{key}'")
    st.session_state[key] = value
    return value


def firewall_get(key, default=None):
    """Read API kept intentionally permissive for compatibility."""
    return st.session_state.get(key, default)


@contextlib.contextmanager
def tab_state_scope(tab_name):
    """Mark which tab is rendering for diagnostics and future enforcement."""
    init_state_firewall()
    previous_scope = st.session_state["_state_firewall"].get("active_scope", "app")
    st.session_state["_state_firewall"]["active_scope"] = tab_name
    try:
        yield
    finally:
        st.session_state["_state_firewall"]["active_scope"] = previous_scope
