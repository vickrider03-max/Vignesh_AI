import json
import os
import time
from datetime import datetime, timedelta

import streamlit as st
from pytz import timezone


CREATOR_USERNAME = "Vignesh"
CREATOR_PASSWORD = "Rider@100"
ACTIVE_USERS_FILE = "active_users.json"


DEFAULT_SESSION_STATE = {
    "is_authenticated": False,
    "logged_in_username": "",
    "user_role": None,
    "login_history": [],
    "uploaded_files": [],
    "selected_files": [],
    "file_texts": {},
    "excel_data_by_file": {},
    "vector_stores": {},
    "messages": [],
    "chat_file_selection": [],
    "chat_summary_downloads": {"images": [], "tables": [], "csv": [], "diagrams": []},
    "start_time": None,
    "user_session_start_time": None,
}


def initialize_session_state():
    try:
        for key, default_value in DEFAULT_SESSION_STATE.items():
            if key not in st.session_state:
                st.session_state[key] = default_value.copy() if isinstance(default_value, (dict, list)) else default_value
    except Exception as exc:
        st.error(f"Error in auth.py at function initialize_session_state: {exc}")


def authenticate_user(username, password):
    try:
        return str(username or "").strip() == CREATOR_USERNAME and str(password or "") == CREATOR_PASSWORD
    except Exception as exc:
        st.error(f"Error in auth.py at function authenticate_user: {exc}")
        return False


def _ist_timestamp():
    try:
        ist_tz = timezone("Asia/Kolkata")
        return datetime.now().astimezone(ist_tz).strftime("%Y-%m-%d %H:%M:%S %Z")
    except Exception as exc:
        st.error(f"Error in auth.py at function _ist_timestamp: {exc}")
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def record_active_user(username):
    try:
        active_users = []
        if os.path.exists(ACTIVE_USERS_FILE):
            with open(ACTIVE_USERS_FILE, "r", encoding="utf-8") as handle:
                active_users = json.load(handle)

        now = datetime.now()
        active_users = [
            user for user in active_users
            if datetime.fromisoformat(user["timestamp"]) > now - timedelta(minutes=30)
            and user["username"] != username
        ]
        active_users.append({"username": username, "timestamp": now.isoformat()})

        with open(ACTIVE_USERS_FILE, "w", encoding="utf-8") as handle:
            json.dump(active_users, handle, indent=2)
    except Exception as exc:
        st.error(f"Error in auth.py at function record_active_user: {exc}")


def clear_active_user(username):
    try:
        if not os.path.exists(ACTIVE_USERS_FILE):
            return
        with open(ACTIVE_USERS_FILE, "r", encoding="utf-8") as handle:
            active_users = json.load(handle)
        active_users = [user for user in active_users if user.get("username") != username]
        with open(ACTIVE_USERS_FILE, "w", encoding="utf-8") as handle:
            json.dump(active_users, handle, indent=2)
    except Exception as exc:
        st.error(f"Error in auth.py at function clear_active_user: {exc}")


def login_user(username, password):
    try:
        initialize_session_state()
        if not authenticate_user(username, password):
            return False

        st.session_state.is_authenticated = True
        st.session_state.logged_in_username = CREATOR_USERNAME
        st.session_state.user_role = "creator"
        st.session_state.start_time = time.time()
        st.session_state.user_session_start_time = _ist_timestamp()
        st.session_state.login_history.append({
            "username": CREATOR_USERNAME,
            "role": "creator",
            "action": "login",
            "timestamp": _ist_timestamp(),
        })
        record_active_user(CREATOR_USERNAME)
        return True
    except Exception as exc:
        st.error(f"Error in auth.py at function login_user: {exc}")
        return False


def logout_user():
    try:
        username = st.session_state.get("logged_in_username", "")
        usage_seconds = int(time.time() - st.session_state.start_time) if st.session_state.get("start_time") else 0
        hours, remainder = divmod(usage_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)

        if username:
            st.session_state.login_history.append({
                "username": username,
                "role": st.session_state.get("user_role"),
                "action": "logout",
                "timestamp": _ist_timestamp(),
                "usage_time": f"{hours}h {minutes}m {seconds}s",
            })
            clear_active_user(username)

        preserved_history = st.session_state.get("login_history", [])
        for key, default_value in DEFAULT_SESSION_STATE.items():
            st.session_state[key] = default_value.copy() if isinstance(default_value, (dict, list)) else default_value
        st.session_state.login_history = preserved_history
        st.rerun()
    except Exception as exc:
        st.error(f"Error in auth.py at function logout_user: {exc}")


def require_authentication():
    try:
        return bool(st.session_state.get("is_authenticated"))
    except Exception as exc:
        st.error(f"Error in auth.py at function require_authentication: {exc}")
        return False
