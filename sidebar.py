import streamlit as st

import auth


SUPPORTED_FILE_TYPES = [
    "pdf", "doc", "docx", "txt", "md", "log", "ppt", "pptx", "xls", "xlsx",
    "csv", "html", "htm", "odt", "rtf", "pages", "capl", "can",
    "png", "jpg", "jpeg", "gif", "bmp", "webp",
]


def _add_uploaded_files(new_files):
    try:
        existing_names = {file_info["name"] for file_info in st.session_state.get("uploaded_files", [])}
        added = False
        for uploaded_file in new_files or []:
            if uploaded_file.name in existing_names:
                continue
            st.session_state.uploaded_files.append({
                "name": uploaded_file.name,
                "bytes": uploaded_file.read(),
            })
            added = True
        if added:
            st.session_state.messages = []
            st.session_state.file_texts = {}
            st.success("New files uploaded.")
    except Exception as exc:
        st.error(f"Error in sidebar.py at function _add_uploaded_files: {exc}")


def _render_file_selector():
    try:
        uploaded_files = st.session_state.get("uploaded_files", [])
        if not uploaded_files:
            st.info("No files uploaded yet.")
            return

        st.markdown("### Uploaded Files")
        current_selection = set(st.session_state.get("selected_files", []))

        for index, file_info in enumerate(uploaded_files):
            file_name = file_info["name"]
            selected = file_name in current_selection
            label = f"Selected: {file_name}" if selected else file_name

            col_file, col_delete = st.columns([0.8, 0.2], vertical_alignment="center")
            with col_file:
                if st.button(label, key=f"sidebar_select_{index}", use_container_width=True):
                    if selected:
                        current_selection.remove(file_name)
                    else:
                        current_selection.add(file_name)
                    st.session_state.selected_files = list(current_selection)
                    st.rerun()
            with col_delete:
                if st.button("Delete", key=f"sidebar_delete_{index}", use_container_width=True):
                    st.session_state.uploaded_files = [
                        item for item in uploaded_files if item["name"] != file_name
                    ]
                    st.session_state.selected_files = [
                        name for name in st.session_state.get("selected_files", []) if name != file_name
                    ]
                    st.session_state.file_texts.pop(file_name, None)
                    st.rerun()
    except Exception as exc:
        st.error(f"Error in sidebar.py at function _render_file_selector: {exc}")


def render_sidebar():
    try:
        with st.sidebar:
            st.markdown("## Workspace")
            st.caption("Upload files, select the ones you need, then choose a workspace.")

            new_files = st.file_uploader(
                "Upload documents",
                type=SUPPORTED_FILE_TYPES,
                accept_multiple_files=True,
            )
            _add_uploaded_files(new_files)

            _render_file_selector()

            if st.button("Clear All Files", use_container_width=True):
                st.session_state.uploaded_files = []
                st.session_state.selected_files = []
                st.session_state.file_texts = {}
                st.session_state.excel_data_by_file = {}
                st.rerun()

            st.divider()
            if st.button("Logout", use_container_width=True):
                auth.logout_user()
    except Exception as exc:
        st.error(f"Error in sidebar.py at function render_sidebar: {exc}")
