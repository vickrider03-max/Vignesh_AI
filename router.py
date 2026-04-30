# ==============================
# STREAMLIT ROUTER ARCHITECTURE
# React-style active tab control with one source of truth.
# Only explicit navigation writes active tab by default.
# ==============================
import streamlit as st

from state_firewall import firewall_set, init_state_firewall
from tab_memory import init_tab_memory


TAB_OPTIONS = ["💬 Chat", "📊 Dashboard", "📂 Compare", "📡 CAPL"]
TAB_KEYS = {
    "💬 Chat": "chat",
    "📊 Dashboard": "dashboard",
    "📂 Compare": "compare",
    "📡 CAPL": "capl",
}


def init_router(default_tab=None):
    """Initialize active-tab state without forcing a rerun."""
    init_state_firewall()
    init_tab_memory()
    default_tab = default_tab or TAB_OPTIONS[0]
    if "active_tab" not in st.session_state:
        firewall_set("app", "active_tab", default_tab)
    if "active_main_tab" not in st.session_state:
        firewall_set("app", "active_main_tab", st.session_state.active_tab)
    if st.session_state.active_main_tab not in TAB_OPTIONS:
        firewall_set("app", "active_main_tab", default_tab)
    if st.session_state.active_tab not in TAB_OPTIONS:
        firewall_set("app", "active_tab", st.session_state.active_main_tab)
    return st.session_state.active_main_tab


def navigate_to(tab_name, explicit=False):
    """Controlled navigation. Implicit switching is ignored unless enabled."""
    init_router()
    if tab_name not in TAB_OPTIONS:
        return st.session_state.active_main_tab
    auto_enabled = st.session_state.get("auto_tab_switch_enabled", False)
    if explicit or auto_enabled:
        firewall_set("app", "active_tab", tab_name)
        firewall_set("app", "active_main_tab", tab_name)
    return st.session_state.active_main_tab


def render_tab_router(label="Open Section"):
    """Render the Streamlit tab radio while preserving existing CSS key names."""
    init_router()
    selected_tab = st.radio(
        label,
        TAB_OPTIONS,
        horizontal=True,
        key="active_main_tab",
        label_visibility="collapsed",
    )
    if selected_tab != st.session_state.get("active_tab"):
        # active_main_tab is owned by the radio widget after instantiation.
        # Only mirror the explicit choice into the router key here.
        firewall_set("app", "active_tab", selected_tab)
    return st.session_state.active_main_tab


def active_tab_key():
    init_router()
    return TAB_KEYS.get(st.session_state.active_main_tab, "chat")
