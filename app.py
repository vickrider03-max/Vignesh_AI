import streamlit as st

import auth
import data_handler
import sidebar
import styles


def render_login_page():
    try:
        styles.render_login_styles()
        st.markdown("## IntelliDoc AI")
        st.caption("Smart Document Assistant")

        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login")

        if submitted:
            if auth.login_user(username, password):
                st.success("Login successful.")
                st.rerun()
            else:
                st.error("Invalid username or password.")
    except Exception as exc:
        st.error(f"Error in app.py at function render_login_page: {exc}")


def render_chat_page():
    try:
        st.subheader("Chat / Analyze")
        selected_files = st.session_state.get("selected_files", [])
        if not selected_files:
            st.info("Select files from the sidebar to analyze them.")
            return

        processed = data_handler.process_selected_files(
            st.session_state.get("uploaded_files", []),
            selected_files,
        )

        action = st.radio(
            "Action",
            ["Analyze document", "Search text", "Show extracted text"],
            horizontal=True,
        )

        if action == "Analyze document":
            for file_name, file_text in processed.items():
                st.markdown(f"### {file_name}")
                st.markdown(data_handler.build_product_documentation_summary(file_name, file_text))

        elif action == "Search text":
            query = st.text_input("Search query")
            if query:
                results = data_handler.search_selected_text(processed, query)
                if results:
                    for file_name, matches in results.items():
                        st.markdown(f"### {file_name}")
                        for line_number, line in matches[:25]:
                            st.markdown(f"- Line {line_number}: `{line}`")
                else:
                    st.info("No matches found.")

        else:
            for file_name, file_text in processed.items():
                with st.expander(file_name, expanded=False):
                    st.text_area(
                        "Extracted text",
                        value=file_text[:20000],
                        height=360,
                        key=f"extracted_text_{file_name}",
                    )
    except Exception as exc:
        st.error(f"Error in app.py at function render_chat_page: {exc}")


def render_dashboard_page():
    try:
        st.subheader("Dashboard")
        uploaded_files = st.session_state.get("uploaded_files", [])
        spreadsheet_files = [
            file_info for file_info in uploaded_files
            if file_info["name"].lower().endswith((".csv", ".xlsx", ".xls"))
        ]

        if not spreadsheet_files:
            st.info("Upload a CSV or Excel file to view tabular data.")
            return

        selected_name = st.selectbox("Spreadsheet", [item["name"] for item in spreadsheet_files])
        file_info = data_handler.get_file_entry(uploaded_files, selected_name)
        if not file_info:
            st.warning("Selected file was not found.")
            return

        df = data_handler.load_tabular_dataframe(selected_name, file_info["bytes"])
        if df.empty:
            st.info("No rows found.")
            return

        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button(
            "Download CSV",
            data=df.to_csv(index=False).encode("utf-8"),
            file_name=f"{selected_name.rsplit('.', 1)[0]}_export.csv",
            mime="text/csv",
        )
    except Exception as exc:
        st.error(f"Error in app.py at function render_dashboard_page: {exc}")


def render_compare_page():
    try:
        st.subheader("Compare")
        selected_files = st.session_state.get("selected_files", [])
        if len(selected_files) < 2:
            st.info("Select at least two files from the sidebar to compare.")
            return

        processed = data_handler.process_selected_files(
            st.session_state.get("uploaded_files", []),
            selected_files[:2],
        )
        names = list(processed.keys())
        comparison = data_handler.compare_texts(processed[names[0]], processed[names[1]])

        st.metric("Similarity", f"{comparison['similarity_percent']}%")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"### {names[0]}")
            st.text_area("Preview A", processed[names[0]][:12000], height=360)
        with col2:
            st.markdown(f"### {names[1]}")
            st.text_area("Preview B", processed[names[1]][:12000], height=360)
    except Exception as exc:
        st.error(f"Error in app.py at function render_compare_page: {exc}")


def render_capl_page():
    try:
        st.subheader("CAPL")
        selected_files = [
            name for name in st.session_state.get("selected_files", [])
            if name.lower().endswith((".can", ".capl", ".txt", ".log"))
        ]
        if not selected_files:
            st.info("Select a CAPL, CAN, TXT, or LOG file from the sidebar.")
            return

        processed = data_handler.process_selected_files(
            st.session_state.get("uploaded_files", []),
            selected_files,
        )
        file_name = st.selectbox("CAPL file", list(processed.keys()))
        st.code(processed[file_name], language="c")
    except Exception as exc:
        st.error(f"Error in app.py at function render_capl_page: {exc}")


def main():
    styles.apply_page_config()
    auth.initialize_session_state()
    styles.apply_global_styles()

    if not auth.require_authentication():
        render_login_page()
        return

    styles.render_header()
    sidebar.render_sidebar()

    page = st.sidebar.radio(
        "Workspace",
        ["Chat", "Dashboard", "Compare", "CAPL"],
        key="active_page",
    )

    routes = {
        "Chat": render_chat_page,
        "Dashboard": render_dashboard_page,
        "Compare": render_compare_page,
        "CAPL": render_capl_page,
    }
    routes[page]()


if __name__ == "__main__":
    main()
