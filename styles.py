import streamlit as st


def apply_page_config():
    try:
        st.set_page_config(page_title="🧠 IntelliDoc AI", layout="wide")
    except Exception as exc:
        st.error(f"Error in styles.py at function apply_page_config: {exc}")


def apply_global_styles():
    try:
        st.markdown(
            """
            <style>
                :root {
                    --surface: #f8fafc;
                    --border: #dbeafe;
                    --text: #173152;
                    --muted: #64748b;
                    --accent: #2563eb;
                }
                .block-container {
                    padding-top: 1.4rem;
                    padding-left: clamp(1rem, 2vw, 2rem);
                    padding-right: clamp(1rem, 2vw, 2rem);
                }
                #MainMenu, footer, header, [data-testid="stToolbar"] {
                    display: none !important;
                }
                .stButton > button,
                div[data-testid="stBaseButton"] button {
                    border-radius: 8px !important;
                    border: 1px solid #bfdbfe !important;
                    background: #eff6ff !important;
                    color: #173152 !important;
                    font-weight: 700 !important;
                }
                .stButton > button:hover {
                    background: #dbeafe !important;
                    border-color: #93c5fd !important;
                }
                [data-testid="stSidebar"] {
                    background: #f8fbff;
                    border-right: 1px solid #dbeafe;
                }
                h1, h2, h3, h4 {
                    color: var(--text);
                }
            </style>
            """,
            unsafe_allow_html=True,
        )
    except Exception as exc:
        st.error(f"Error in styles.py at function apply_global_styles: {exc}")


def render_login_styles():
    try:
        st.markdown(
            """
            <style>
                [data-testid="stAppViewContainer"] {
                    background: linear-gradient(135deg, #ffffff 0%, #f8fafc 45%, #eaf6ff 100%);
                }
                .block-container {
                    max-width: 520px;
                    padding-top: 15vh;
                }
            </style>
            """,
            unsafe_allow_html=True,
        )
    except Exception as exc:
        st.error(f"Error in styles.py at function render_login_styles: {exc}")


def render_header():
    try:
        col_title, col_user = st.columns([5, 2], vertical_alignment="center")
        with col_title:
            st.markdown("## IntelliDoc AI")
            st.caption("Smart Document Assistant")
        with col_user:
            user = st.session_state.get("logged_in_username", "User")
            role = st.session_state.get("user_role", "user")
            st.markdown(f"**{user}**")
            st.caption(str(role).title())
        st.divider()
    except Exception as exc:
        st.error(f"Error in styles.py at function render_header: {exc}")
