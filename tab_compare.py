# Auto-generated from legacy_app.py during modular refactor.
# The original monolith is retained as rollback documentation.

from functions import *

# ==============================
# COMPARE TAB UI
# Moved from legacy_app.py tab body.
# UI rendering and event handling live here; backend work is delegated to functions.py.
# ==============================
def render_compare_tab():
    st.markdown('<div id="compare-section">', unsafe_allow_html=True)
    compare_header_col, compare_reset_col = st.columns([8, 1])
    with compare_header_col:
        st.subheader("Compare Files")
    with compare_reset_col:
        if st.button("🧼 Reset", key="reset_compare_selection"):
            st.session_state.compare_file_selection = []
            st.session_state.compare_result_html = None
            st.session_state.compare_result_excel_bytes = None
            st.session_state.compare_result_files = []
            st.session_state.compare_semantic_summary = None
            st.rerun()

    st.info("Select files in the sidebar to make them available here, then choose only the files you want to compare in this tab.")
    show_current_sidebar_selection()
    show_help_popup('compare', st.session_state.selected_files)
    render_file_context_card("Compare File Context", st.session_state.selected_files, st.session_state.compare_file_selection)

    st.markdown("**Comparison options:**")
    st.markdown(
        "- Inline word-level diff (sequence-aware highlighting)\n"
        "- Side-by-side line diff for direct file comparison\n"
        "- Word presence summary across selected files\n"
        "- Downloadable Excel comparison workbook\n"
        "- Supports PDF, DOCX, PPTX, XLSX, HTML, TXT, CAN/CAPL formats"
    )

    # Use selected files in multiselect (user must choose from selected_files independently)
    st.session_state.compare_file_selection = [
        file_name for file_name in st.session_state.compare_file_selection
        if file_name in st.session_state.selected_files
    ]
    selected_files_for_comparison = st.multiselect(
        "Choose files to compare",
        options=st.session_state.selected_files,
        default=st.session_state.compare_file_selection,
        key="compare_file_selection"
    )

    compare_mode = st.selectbox(
        "Comparison mode",
        ["Exact inline word diff", "Side-by-side line diff", "Word presence summary"],
        index=0,
        key="compare_mode"
    )

    reference_file = None
    # Reference file selection is no longer shown; the first selected file is used automatically for comparison baselines.
    if len(selected_files_for_comparison) == 1:
        st.info("Select at least two files to compare.")

    compare_clicked = st.button("Compare Selected Files", key="run_compare_button", use_container_width=True)

    if compare_clicked:
        if len(selected_files_for_comparison) >= 2:
            with st.spinner("Loading files for comparison..."):
                ensure_files_processed(selected_files_for_comparison)

            selected_texts = {}
            for f in selected_files_for_comparison:
                raw_text = st.session_state.file_texts.get(f, "")
                selected_texts[f] = raw_text if isinstance(raw_text, str) else str(raw_text)

            with st.spinner("Loading comparison results..."):
                html_diff = highlight_multi_file_differences(
                    selected_texts,
                    comparison_mode=compare_mode,
                    reference_file=reference_file
                )
                excel_io = generate_word_level_comparison_excel(selected_texts)
                semantic_summary = build_semantic_diff_explanation(selected_texts)

            st.session_state.compare_result_html = html_diff
            st.session_state.compare_result_excel_bytes = excel_io.getvalue()
            st.session_state.compare_result_files = selected_files_for_comparison.copy()
            st.session_state.compare_semantic_summary = semantic_summary
            record_workspace_memory_event(
                "compare",
                "Semantic comparison completed",
                semantic_summary,
                source=", ".join(selected_files_for_comparison),
            )
            save_workspace_memory()
            save_memory_log("compare", "Semantic comparison stored in workspace memory.", {
                "files": selected_files_for_comparison,
            })
        else:
            st.warning("Select at least two files to compare.")

    if st.session_state.compare_result_html and st.session_state.compare_result_files:
        st.info("Compared files: " + ", ".join(st.session_state.compare_result_files))
        st.markdown(f"### Comparison Results ({len(st.session_state.compare_result_files)} files)")
        if st.session_state.get("compare_semantic_summary"):
            st.markdown(st.session_state.compare_semantic_summary, unsafe_allow_html=True)
        render_html_frame(st.session_state.compare_result_html, height=800)

        st.markdown("### Download Excel Comparison")
        st.download_button(
            "Download Comparison Excel",
            st.session_state.compare_result_excel_bytes,
            file_name="comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    elif len(selected_files_for_comparison) < 2:
        st.info("Select at least two files to compare.")

    st.markdown('</div>', unsafe_allow_html=True)

    # -------------------------------
