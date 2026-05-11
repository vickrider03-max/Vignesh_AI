```python
# ==============================
# CHAT TAB UI
# Optimized + cleaned merged version
# ==============================

from functions import *
from tabs.tab_memory import get_tab_uploaded_files

DOCUMENT_INTELLIGENCE_INTENTS = {
    "FULL_DOCUMENT_ANALYSIS",
    "SHORT_SUMMARY",
    "OVERVIEW",
    "FEATURES_ONLY",
    "PIN_DIAGRAMS_CONNECTORS_TABLES",
    "WORKFLOW_OR_PROCESS",
    "USE_CASES_APPLICATIONS",
    "TABLE_EXTRACTION",
    "IMAGE_OR_DIAGRAM_EXPLANATION",
    "DOWNLOADABLE_REPORT",
    "SPECIFIC_COMPONENT_DETAILS",
    "DEVICE_EQUIPMENT_EXPLANATION",
    "EXTRACTION",
}

def render_chat_tab():

    st.markdown('<div id="chat-section">', unsafe_allow_html=True)

    # -------------------------------------------------
    # SESSION STATE INIT
    # -------------------------------------------------

    defaults = {
        "input_prefill": "",
        "chat_next_suggestions": [],
        "chat_next_suggestions_for": None,
        "document_chat_display": {},
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    # -------------------------------------------------
    # HELPERS
    # -------------------------------------------------

    def normalize_chat_quick_action(text):
        clean_text = str(text or "").strip()
        lower_text = clean_text.lower()

        mappings = {
            "analyze data": "analyze",
            "overview": "overview",
            "find keyword": 'find "phrase"',
            "find phrase": 'find "phrase"',
            "count signals": 'count "signal"',
        }

        return mappings.get(lower_text, clean_text)

    def normalize_chat_slash_command(text):

        clean_text = normalize_chat_quick_action(text)

        if not clean_text.startswith("/"):
            return clean_text

        command, _, remainder = clean_text.partition(" ")

        command = command.lower().strip()
        remainder = remainder.strip()

        mapping = {
            "/analyze": f"analyze {remainder}",
            "/compare": f"compare {remainder}",
            "/overview": f"overview {remainder}",
        }

        if command == "/find":
            if remainder and not re.search(r"'(.*?)'|\"(.*?)\"", remainder):
                return f'find "{remainder}"'
            return f"find {remainder}".strip()

        return mapping.get(command, clean_text)

    def get_valid_chat_message(message):

        if not isinstance(message, dict):
            return None

        role = str(message.get("role", "")).strip().lower()

        if role not in {"user", "assistant"}:
            return None

        content = str(message.get("content", "")).strip()

        if not content:
            return None

        return {
            "role": role,
            "content": content
        }

    def enforce_document_intelligence_output_rules(response_text):

        text = str(response_text or "").strip()

        if not text:
            return text

        text = re.sub(r"(?im)^\s*Answer:\s*\n?", "", text)

        text = re.sub(
            r"(?im)^confidence:\s*(very\s+high|high|medium|moderate|low|very\s+low)\s*$",
            "",
            text,
        )

        text = re.sub(
            r"(?i)\bNot specified in the provided context\.?",
            MISSING_DOCUMENT_INFO_MESSAGE,
            text,
        )

        return text.strip()

    # -------------------------------------------------
    # HEADER
    # -------------------------------------------------

    chat_header_col, chat_reset_col = st.columns([8, 1])

    with chat_header_col:
        st.subheader("Chat with Selected Documents")

    with chat_reset_col:

        if st.button("🧼 Reset", key="reset_chat_selection"):

            st.session_state.chat_file_selection = []
            st.session_state.messages = []
            st.session_state.document_chat_display = {}

            st.success("✅ Chat reset!")
            st.rerun()

    # -------------------------------------------------
    # FILE SELECTION
    # -------------------------------------------------

    available_chat_files = list(dict.fromkeys(st.session_state.selected_files))

    if not available_chat_files:
        st.info("Select files from sidebar to start chatting.")
        return

    st.session_state.chat_file_selection = [
        file_name
        for file_name in st.session_state.get("chat_file_selection", [])
        if file_name in available_chat_files
    ]

    chat_files = st.multiselect(
        "Choose file(s) for Chat",
        options=available_chat_files,
        default=st.session_state.chat_file_selection,
        key="chat_file_selection"
    )

    if not chat_files:
        st.info("Choose one or more files.")
        return

    # -------------------------------------------------
    # LOAD FILES
    # -------------------------------------------------

    with st.spinner("Loading files..."):
        ensure_files_processed(chat_files)

    selected_file_texts = {
        f: st.session_state.file_texts.get(f, "")
        for f in chat_files
    }

    user_id = get_active_user_id()

    chat_display_key = get_chatpdf_memory_key(user_id, chat_files)

    current_chat_messages = st.session_state.document_chat_display.setdefault(
        chat_display_key,
        []
    )

    # -------------------------------------------------
    # CHAT INPUT
    # -------------------------------------------------

    user_input = st.chat_input("Ask anything related to selected documents")

    if st.session_state.get("input_prefill"):
        user_input = st.session_state.input_prefill
        st.session_state.input_prefill = ""

    # -------------------------------------------------
    # PROCESS INPUT
    # -------------------------------------------------

    if user_input:

        submitted_input = user_input

        processing_input = normalize_chat_slash_command(user_input)

        current_chat_messages.append({
            "role": "user",
            "content": submitted_input
        })

        st.session_state.messages = current_chat_messages

        with st.spinner("Analyzing document..."):

            user_input_lower = processing_input.lower()

            technical_request_type = classify_technical_document_request(
                processing_input
            )

            citation_docs = []

            response = ""

            # -------------------------------------------------
            # COUNT QUERY
            # -------------------------------------------------

            is_count_query = any(
                t in user_input_lower
                for t in ["how many", "count", "number of", "occurrences"]
            )

            if is_count_query:

                match = re.search(r"'(.*?)'|\"(.*?)\"", processing_input)

                if match:

                    word = match.group(1) or match.group(2)

                    total_count = 0

                    for file_text in selected_file_texts.values():

                        total_count += len(
                            re.findall(
                                rf'(?<![\w-]){re.escape(word)}(?![\w-])',
                                file_text,
                                re.IGNORECASE
                            )
                        )

                    response = (
                        f"🔢 '{word}' appears "
                        f"{total_count} times "
                        f"in selected documents."
                    )

                else:

                    response = (
                        "⚠️ Specify phrase in quotes.\n\n"
                        "Example:\n"
                        'count("CAN")'
                    )

            # -------------------------------------------------
            # FIND QUERY
            # -------------------------------------------------

            elif any(
                term in user_input_lower
                for term in ["find", "search", "locate"]
            ):

                match = re.search(r"'(.*?)'|\"(.*?)\"", processing_input)

                if match:

                    query = match.group(1) or match.group(2)

                    response_blocks = []

                    for f in chat_files:

                        file_text = st.session_state.file_texts.get(f, "")

                        response_blocks.append(
                            build_highlighted_search_results(
                                f,
                                file_text,
                                query
                            )
                        )

                    response = "\n".join(response_blocks)

                else:

                    response = (
                        "⚠️ Specify search term in quotes.\n\n"
                        'find("VT6104B")'
                    )

            # -------------------------------------------------
            # DOCUMENT INTELLIGENCE
            # -------------------------------------------------

            elif technical_request_type in DOCUMENT_INTELLIGENCE_INTENTS:

                response, citation_docs = smart_file_brain_query(
                    processing_input,
                    chat_files,
                    user_id=user_id,
                    intent=technical_request_type,
                    top_k=15
                )

            # -------------------------------------------------
            # COMPARISON
            # -------------------------------------------------

            elif technical_request_type == "COMPARISON":

                compared_items = extract_multiple_component_names(
                    processing_input
                )

                if len(compared_items) >= 2:

                    response = build_component_comparison_response(
                        selected_file_texts,
                        processing_input
                    )

                else:

                    response = (
                        "⚠️ Mention at least 2 components to compare."
                    )

            # -------------------------------------------------
            # FALLBACK QA
            # -------------------------------------------------

            else:

                response, citation_docs = answer_chatpdf_question(
                    processing_input,
                    chat_files,
                    user_id=user_id,
                    top_k=15,
                )

            # -------------------------------------------------
            # EMPTY RESPONSE SAFETY
            # -------------------------------------------------

            response = str(response or "").strip()

            if not response:

                response = (
                    "Unable to generate response.\n"
                    "Try another query."
                )

            # -------------------------------------------------
            # SOURCE APPEND
            # -------------------------------------------------

            if citation_docs:

                sources_text = format_chatpdf_sources(citation_docs)

                response += "\n\nSources:\n" + sources_text

            # -------------------------------------------------
            # CLEANUP
            # -------------------------------------------------

            response = enforce_document_intelligence_output_rules(
                response
            )

            # -------------------------------------------------
            # SAVE CHAT
            # -------------------------------------------------

            current_chat_messages.append({
                "role": "assistant",
                "content": response
            })

            st.session_state.document_chat_display[
                chat_display_key
            ] = current_chat_messages[-100:]

            st.session_state.messages = (
                st.session_state.document_chat_display[
                    chat_display_key
                ]
            )

    # -------------------------------------------------
    # RENDER CHAT
    # -------------------------------------------------

    visible_messages = st.session_state.get("messages", [])

    for msg_index, raw_msg in enumerate(visible_messages):

        msg = get_valid_chat_message(raw_msg)

        if msg is None:
            continue

        avatar = "🧑" if msg["role"] == "user" else "🤖"

        with st.chat_message(msg["role"], avatar=avatar):

            st.markdown(msg["content"])

    st.markdown('</div>', unsafe_allow_html=True)
```
