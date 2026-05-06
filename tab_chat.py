# Auto-generated from legacy_app.py during modular refactor.
# The original monolith is retained as rollback documentation.

from functions import *
from tab_memory import get_tab_uploaded_files

# ==============================
# CHAT TAB UI
# Moved from legacy_app.py tab body.
# UI rendering and event handling live here; backend work is delegated to functions.py.
# ==============================
def render_chat_tab():
    st.markdown('<div id="chat-section">', unsafe_allow_html=True)
    st.markdown(
        """
        <style>
        [class*="st-key-chat_sugg_"] button,
        [class*="st-key-ai_sugg_"] button {
            min-height: 38px !important;
            border-radius: 999px !important;
            border: 1px solid rgba(147, 197, 253, 0.52) !important;
            background: rgba(248, 251, 255, 0.88) !important;
            color: #173152 !important;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06) !important;
            font-size: 0.88rem !important;
            font-weight: 700 !important;
            padding: 0.42rem 0.75rem !important;
            transition: transform 0.18s ease, border-color 0.18s ease, background-color 0.18s ease !important;
            white-space: nowrap !important;
        }
        [class*="st-key-chat_sugg_"] button:hover,
        [class*="st-key-ai_sugg_"] button:hover {
            background: rgba(239, 246, 255, 0.98) !important;
            border-color: rgba(59, 130, 246, 0.55) !important;
            transform: translateY(-1px) scale(1.015) !important;
        }
        .chat-ghost-hint {
            color: #2563eb;
            font-size: 0.86rem;
            margin: 0.1rem 0 0.45rem;
        }
        .st-key-chat_file_selection div[data-baseweb="select"] > div {
            background: #f8fbff !important;
            border: 1px solid #cfe2f3 !important;
            border-radius: 12px !important;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.04) !important;
            min-height: 44px !important;
        }
        .st-key-chat_file_selection div[data-baseweb="select"]:focus-within > div {
            border-color: #60a5fa !important;
            box-shadow: 0 0 0 3px rgba(96, 165, 250, 0.16) !important;
        }
        .st-key-chat_file_selection span[data-baseweb="tag"] {
            background: #e8f6ff !important;
            border: 1px solid #c0dff0 !important;
            border-radius: 10px !important;
            color: #173152 !important;
            font-weight: 650 !important;
        }
        [data-testid="stChatInput"] {
            margin-top: 1rem !important;
        }
        [data-testid="stChatInput"] > div {
            background: #ffffff !important;
            border: 1px solid #cfe2f3 !important;
            border-radius: 16px !important;
            box-shadow: 0 10px 28px rgba(15, 23, 42, 0.07) !important;
            padding: 3px 6px !important;
        }
        [data-testid="stChatInput"] > div:focus-within {
            border-color: #60a5fa !important;
            box-shadow: 0 0 0 3px rgba(96, 165, 250, 0.15), 0 12px 30px rgba(15, 23, 42, 0.08) !important;
        }
        [data-testid="stChatInput"] textarea {
            background: transparent !important;
            color: #173152 !important;
            caret-color: #2563eb !important;
            min-height: 44px !important;
            padding: 0.75rem 0.9rem !important;
            font-size: 0.98rem !important;
            font-weight: 500 !important;
            outline: none !important;
            box-shadow: none !important;
        }
        [data-testid="stChatInput"] textarea::placeholder {
            color: #7b8aa0 !important;
            opacity: 1 !important;
        }
        [data-testid="stChatInput"] textarea:focus,
        [data-testid="stChatInput"] textarea:focus-visible {
            outline: none !important;
            box-shadow: none !important;
            border-color: transparent !important;
        }
        [data-testid="stChatInput"] div[data-baseweb="base-input"],
        [data-testid="stChatInput"] div[data-baseweb="textarea"] {
            background: transparent !important;
            border: 0 !important;
            box-shadow: none !important;
        }
        [data-testid="stChatInput"] button {
            background: #e8f6ff !important;
            border: 1px solid #c0dff0 !important;
            border-radius: 12px !important;
            color: #173152 !important;
            height: 38px !important;
            width: 38px !important;
            min-height: 38px !important;
            min-width: 38px !important;
            margin-right: 2px !important;
            box-shadow: none !important;
        }
        [data-testid="stChatInput"] button:hover {
            background: #dbeafe !important;
            border-color: #93c5fd !important;
            transform: translateY(-1px) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    if "input_prefill" not in st.session_state:
        st.session_state.input_prefill = ""
    if "chat_next_suggestions" not in st.session_state:
        st.session_state.chat_next_suggestions = []
    if "chat_next_suggestions_for" not in st.session_state:
        st.session_state.chat_next_suggestions_for = None

    def get_chat_hint(text):
        text = str(text or "").lower()
        if "vn1630a" in text:
            return "-> fetching component details"
        if "d-sub9" in text:
            return "-> generating pin diagram"
        if "count" in text:
            return "-> analyzing signals"
        if "find" in text or "search" in text or "locate" in text:
            return "-> searching knowledge base"
        return ""

    def extract_chat_entity(text):
        quoted = extract_quoted_item_name(text)
        if quoted:
            return quoted
        match = re.search(r"\b[A-Z]{2,}[A-Z0-9_-]{2,}\b", str(text or ""))
        return match.group(0) if match else ""

    def normalize_chat_quick_action(text):
        clean_text = str(text or "").strip()
        lower_text = clean_text.lower()
        if lower_text == "analyze data":
            return "analyze"
        if lower_text == "overview":
            return "overview"
        if lower_text == "find keyword":
            return 'find "keyword"'
        if lower_text == "count signals":
            return 'count "signal"'
        if lower_text.startswith("item details:"):
            item = clean_text.split(":", 1)[1].strip()
            return f'item details "{item}"' if item else clean_text
        if lower_text.startswith("pin diagram:"):
            item = clean_text.split(":", 1)[1].strip()
            return f'pin diagram "{item}"' if item else clean_text
        return clean_text

    def normalize_chat_slash_command(text):
        clean_text = normalize_chat_quick_action(text)
        if not clean_text.startswith("/"):
            return clean_text
        command, _, remainder = clean_text.partition(" ")
        command = command.lower().strip()
        remainder = remainder.strip()
        if command == "/analyze":
            return f"analyze {remainder}".strip()
        if command == "/compare":
            return f"compare {remainder}".strip()
        if command == "/overview":
            return f"overview {remainder}".strip()
        if command == "/find":
            if remainder and not re.search(r"'(.*?)'|\"(.*?)\"", remainder):
                return f'find "{remainder}"'
            return f"find {remainder}".strip()
        return clean_text

    def generate_chat_reasoning(user_input, context):
        text = str(user_input or "")
        combined = f"{text}\n{str(context or '')[:6000]}".lower()
        reasoning = [
            "Identify entity",
            "Retrieve memory context",
            "Analyze intent",
            "Detect missing information",
            "Find next actions",
        ]
        if any(term in combined for term in ["diagram", "pin", "d-sub", "connector"]):
            reasoning.append("Diagram requested or available")
        if any(term in combined for term in ["compare", "difference", "diff"]):
            reasoning.append("Compare related context")
        if any(term in combined for term in ["signal", "signals", "count"]):
            reasoning.append("Signal analysis needed")
        return list(dict.fromkeys(reasoning))

    def build_chat_next_suggestions(user_input, context, intent=None):
        if not should_show_chat_suggestions(intent or classify_document_chat_intent(user_input), user_input):
            return []

        suggestions = []
        memory_hits = search_workspace_memory(user_input, limit=3)
        for memory_item in memory_hits:
            memory_text = str(memory_item or "").lower()
            if "diagram" in memory_text:
                suggestions.append("Show pin diagram")
            if "signal" in memory_text:
                suggestions.append("Count signals")
            if "compare" in memory_text or "difference" in memory_text:
                suggestions.append("Compare with similar items")
            if "overview" in memory_text or "summary" in memory_text:
                suggestions.append("Get document overview")
            if ("table" in memory_text or "data" in memory_text) and intent in {"ANALYSIS", "SUMMARY", "GUIDANCE"}:
                suggestions.append("Inspect relevant tables")

        entity = extract_chat_entity(user_input)
        if entity:
            if any(term in user_input.lower() for term in ["pin", "diagram", "connector", "d-sub"]):
                suggestions.insert(0, f"Pin diagram: {entity}")
            else:
                suggestions.insert(0, f"Item details: {entity}")

        return list(dict.fromkeys(suggestions))[:4]

    def get_valid_chat_message(message):
        """Return a normalized chat message or None for legacy/empty entries."""
        if not isinstance(message, dict):
            return None
        role = str(message.get("role", "")).strip().lower()
        if role not in {"user", "assistant"}:
            return None
        content = str(message.get("content", "") or "").strip()
        if not content:
            return None
        return {"role": role, "content": content}

    def render_chat_message_content(message, msg_index):
        content = message["content"]
        if message["role"] == "assistant" and msg_index == st.session_state.get("last_streamed_assistant_index"):
            placeholder = st.empty()
            tokens = re.split(r"(\s+)", content)
            streamed = ""
            for token_index, token in enumerate(tokens):
                streamed += token
                if token_index < 240:
                    placeholder.markdown(streamed + "▌", unsafe_allow_html=True)
                    time.sleep(0.006)
            placeholder.markdown(content, unsafe_allow_html=True)
            st.session_state.last_streamed_assistant_index = None
        elif message["role"] == "assistant":
            st.markdown(content, unsafe_allow_html=True)
        else:
            st.markdown(content)

    chat_header_col, chat_reset_col = st.columns([8, 1])
    with chat_header_col:
        st.subheader("Chat with Selected Documents")
    with chat_reset_col:
        if st.button(" 🧼 Reset", key="reset_chat_selection", help="Reset chat selection"):
            st.session_state.chat_file_selection = []
            st.session_state.chat_summary_downloads = empty_chat_summary_downloads()
            st.session_state.messages = []
            st.session_state.input_prefill = ""
            st.session_state.chat_next_suggestions = []
            st.session_state.chat_next_suggestions_for = None
            st.success("✅ Chat reset!")
            st.rerun()

    st.info(
        "Choose files in the sidebar to make them available here. Then select only the files you want for Chat in this tab.")
    chat_tab_file_names = [file_dict.get("name") for file_dict in get_tab_uploaded_files("chat") if file_dict.get("name")]
    available_chat_files = list(dict.fromkeys(chat_tab_file_names + st.session_state.selected_files))

    show_current_sidebar_selection()
    render_file_context_card("Chat File Context", available_chat_files, st.session_state.chat_file_selection)

    show_help_popup('chat', available_chat_files)

    if available_chat_files:
        st.session_state.chat_file_selection = [
            file_name for file_name in st.session_state.chat_file_selection
            if file_name in available_chat_files
        ]
        chat_files = st.multiselect("Choose file(s) for Chat", options=available_chat_files,
                                    default=st.session_state.chat_file_selection, key="chat_file_selection")
        if not chat_files:
            st.info("Choose one or more files in this tab to start chatting.")
        else:
            with st.spinner("Loading selected files..."):
                ensure_files_processed(chat_files)
            selected_file_texts = {f: st.session_state.file_texts.get(f, "") for f in chat_files}
            combined_text = "\n".join(selected_file_texts.values())
            user_id = get_active_user_id()
            document_memory = get_chatpdf_memory(user_id, chat_files)
            chat_display_key = get_chatpdf_memory_key(user_id, chat_files)
            if "document_chat_display" not in st.session_state or not isinstance(st.session_state.document_chat_display, dict):
                st.session_state.document_chat_display = {}
            current_chat_messages = st.session_state.document_chat_display.setdefault(chat_display_key, [])
            
            # Check if files are actually loaded
            if not combined_text or not any(selected_file_texts.values()):
                st.warning("⚠️ Files are selected but their content is not loaded yet. Please wait a moment and try again, or re-select the files.")

            st.caption(
                f"ChatPDF mode: {len(chat_files)} document(s) selected, "
                f"{len(document_memory)} memory entr{'y' if len(document_memory) == 1 else 'ies'} for this selection."
            )

            user_input = st.chat_input("Ask anything related to selected documents/files")
            if st.session_state.get("input_prefill"):
                user_input = st.session_state.input_prefill
                st.session_state.input_prefill = ""
            if user_input:
                submitted_input = user_input
                processing_input = normalize_chat_slash_command(user_input)
                hint = get_chat_hint(processing_input or submitted_input)
                if hint:
                    st.markdown(f"<div class='chat-ghost-hint'>{html.escape(hint)}</div>", unsafe_allow_html=True)

                if submitted_input.strip().lower() == "clear":
                    current_chat_messages.clear()
                    st.session_state.messages = []
                    st.session_state.chat_summary_downloads = empty_chat_summary_downloads()
                    st.session_state.chat_next_suggestions = []
                    st.session_state.chat_next_suggestions_for = None
                    st.success("✅ Chat cleared!")
                else:
                    current_chat_messages.append({"role": "user", "content": submitted_input})
                    st.session_state.messages = current_chat_messages
                    with st.spinner("Processing your request..."):
                        st.session_state.chat_summary_downloads = empty_chat_summary_downloads()
                        user_input_lower = processing_input.lower()
                        chat_intent = classify_document_chat_intent(processing_input)
                        technical_request_type = classify_technical_document_request(processing_input)
                        document_profile = detect_document_chat_profile(chat_files, combined_text)
                        is_count_query = any(t in user_input_lower for t in ["how many", "count", "number of", "occurrences"])
                        is_find_query = any(term in user_input_lower for term in ["find", "search", "locate"]) or "highlight" in user_input_lower
                        citation_docs = []
                        explicit_full_analysis = bool(re.search(r"\b(analy[sz]e|analysis|full analysis|detailed analysis|complete analysis|analyze document|analyse document|explain document)\b", user_input_lower))
                        explicit_summary = bool(re.search(r"\b(summary|summari[sz]e|short summary|brief summary|main points|key points|recap)\b", user_input_lower))
                        if explicit_full_analysis:
                            technical_request_type = "FULL_DOCUMENT_ANALYSIS"
                        elif explicit_summary:
                            technical_request_type = "SHORT_SUMMARY"
                        explicit_document_action = (
                            explicit_full_analysis
                            or explicit_summary
                            or technical_request_type in {
                                "OVERVIEW",
                                "FEATURES_ONLY",
                                "PIN_DIAGRAMS_CONNECTORS_TABLES",
                                "WORKFLOW_OR_PROCESS",
                                "USE_CASES_APPLICATIONS",
                                "TABLE_EXTRACTION",
                                "IMAGE_OR_DIAGRAM_EXPLANATION",
                                "DOWNLOADABLE_REPORT",
                                "SPECIFIC_COMPONENT_DETAILS",
                                "COMPARISON",
                                "TROUBLESHOOTING_OR_LIMITATIONS",
                                "REQUIREMENTS_OR_SPECIFICATION_EXTRACTION",
                            }
                        )
                        # Word count queries
                        if is_count_query:
                            match = re.search(r"'(.*?)'|\"(.*?)\"", processing_input)
                            if match:
                                word = match.group(1) or match.group(2)
                                count = len(
                                    re.findall(rf'(?<![\w-]){re.escape(word)}(?![\w-])', combined_text, re.IGNORECASE))
                                response = f"🔢 The word/phrase '{word}' appears {count} times in the selected documents."
                            elif "vn" in user_input_lower and any(term in user_input_lower for term in ["device", "devices", "interface", "module", "modules"]):
                                extracted_response = build_extraction_response_for_query(processing_input, selected_file_texts)
                                device_count = len(list(dict.fromkeys(extract_vn_devices_from_text(combined_text))))
                                response = f"**VN device count:** {device_count}\n\n{extracted_response}"
                            else:
                                response = "⚠️ Specify the word/phrase in quotes. Example: count('keyword') or count(\"keyword\")"
                        elif is_find_query:
                            match = re.search(r"'(.*?)'|\"(.*?)\"", processing_input)
                            if match:
                                query = match.group(1) or match.group(2)
                                response_blocks = []
                                for f in chat_files:
                                    file_text = st.session_state.file_texts.get(f, "")
                                    response_blocks.append(build_highlighted_search_results(f, file_text, query))
                                response = "".join(response_blocks)
                            else:
                                response = "⚠️ Specify the search word or phrase in quotes. Example: find('keyword') or search(\"keyword\")"
                        elif technical_request_type == "FULL_DOCUMENT_ANALYSIS":
                            response = build_full_document_summary_response(selected_file_texts)
                        elif technical_request_type == "SHORT_SUMMARY":
                            response = build_short_summary_response(selected_file_texts)
                        elif technical_request_type == "OVERVIEW":
                            response = build_overview_response(selected_file_texts)
                        elif technical_request_type == "FEATURES_ONLY":
                            response = build_features_only_response(selected_file_texts)
                        elif technical_request_type == "PIN_DIAGRAMS_CONNECTORS_TABLES":
                            response, pin_csv_downloads, ascii_diagram_downloads = build_diagram_pin_details_response(selected_file_texts, processing_input)
                            st.session_state.chat_summary_downloads = {
                                "images": [],
                                "tables": [],
                                "csv": pin_csv_downloads,
                                "diagrams": ascii_diagram_downloads,
                            }
                        elif technical_request_type == "WORKFLOW_OR_PROCESS":
                            response = build_workflow_or_process_response(selected_file_texts)
                        elif technical_request_type == "USE_CASES_APPLICATIONS":
                            response = build_use_cases_applications_response(selected_file_texts)
                        elif technical_request_type == "TABLE_EXTRACTION":
                            response = build_table_extraction_response(selected_file_texts)
                        elif technical_request_type == "IMAGE_OR_DIAGRAM_EXPLANATION":
                            response = build_image_or_diagram_extraction_response(selected_file_texts, processing_input)
                        elif technical_request_type == "DOWNLOADABLE_REPORT": 
                            response = build_downloadable_report_response(selected_file_texts)
                        elif technical_request_type == "SPECIFIC_COMPONENT_DETAILS":
                            response = build_specific_component_response(selected_file_texts, processing_input)
                        elif technical_request_type == "COMPARISON":
                            compared_items = extract_multiple_component_names(processing_input)
                            if len(compared_items) >= 2:
                                response = build_component_comparison_response(selected_file_texts, processing_input)
                            elif len(chat_files) >= 2:
                                selected_texts = {f: st.session_state.file_texts[f] for f in chat_files}
                                response = highlight_multi_file_differences(selected_texts)
                            else:
                                response = "⚠️ Please mention two items/components or select at least 2 files to compare."
                        elif technical_request_type == "TROUBLESHOOTING_OR_LIMITATIONS":
                            response = build_troubleshooting_or_limitations_response(selected_file_texts)
                        elif technical_request_type == "REQUIREMENTS_OR_SPECIFICATION_EXTRACTION":
                            response = build_requirements_or_specification_extraction_response(selected_file_texts)
                        elif chat_intent == "EXTRACTION":
                            response = build_extraction_response_for_query(processing_input, selected_file_texts)
                        elif not explicit_document_action:
                            response, citation_docs = answer_chatpdf_question(
                                processing_input,
                                chat_files,
                                user_id=user_id,
                                top_k=7,
                            )
                        elif chat_intent == "UNKNOWN":
                            response = build_short_summary_response(selected_file_texts)
                        else:
                            combined_vs = get_workspace_vector_store(chat_files) or get_combined_vector_store(chat_files)
                            retriever = combined_vs.as_retriever(search_kwargs={"k": 3})
                            llm = load_llm()
                            chat_history = "\n".join(
                                f"{'User' if msg['role'] == 'user' else 'Assistant'}: {msg['content']}"
                                for msg in st.session_state.messages[:-1]
                            )
                            prompt = ChatPromptTemplate.from_messages([
                                ("system",
                                 "You are an Enterprise Document Intelligence Engine.\n\n"
                                 "You analyze ANY uploaded document type, including PDF, DOCX, DOC, PPTX, PPT, XLSX, XLS, CSV, TXT, HTML, Markdown, RTF, ODT, images, technical manuals, reports, specifications, presentations, spreadsheets, and mixed-format documents.\n\n"
                                 "You must answer based on the USER'S EXACT REQUEST, not by summarizing the entire document every time.\n\n"
                                 "DOCUMENT CONTEXT:\n"
                                 "The provided document content may contain OCR noise, metadata, cover pages, copyright pages, imprint text, table of contents, page numbers, headers, footers, repeated section titles, broken words, incomplete lines, extracted fragments, images, diagrams, tables, or partial sections.\n\n"
                                 "Your job is to filter noise, identify meaningful content, and produce a professional, human-readable answer.\n\n"
                                 "1. Internal intent detection\n\n"
                                 "First classify the user request internally into ONE primary intent:\n\n"
                                 "FULL_DOCUMENT_ANALYSIS\n"
                                 "SHORT_SUMMARY\n"
                                 "OVERVIEW\n"
                                 "FEATURES_ONLY\n"
                                 "SPECIFIC_COMPONENT_DETAILS\n"
                                 "PIN_DIAGRAMS_CONNECTORS_TABLES\n"
                                 "WORKFLOW_OR_PROCESS\n"
                                 "USE_CASES_APPLICATIONS\n"
                                 "COMPARISON\n"
                                 "TABLE_EXTRACTION\n"
                                 "IMAGE_OR_DIAGRAM_EXPLANATION\n"
                                 "DOWNLOADABLE_REPORT\n"
                                 "TROUBLESHOOTING_OR_LIMITATIONS\n"
                                 "REQUIREMENTS_OR_SPECIFICATION_EXTRACTION\n\n"
                                 "Do not display this classification unless the user asks.\n\n"
                                 "If the user request contains multiple intents, satisfy them in priority order and avoid repeating the same information.\n\n"
                                 "2. Context quality check before answering\n\n"
                                 "Before answering, evaluate whether the provided context contains enough meaningful content.\n\n"
                                 "Low-quality context includes mostly:\n\n"
                                 "Metadata\n"
                                 "Cover page text\n"
                                 "Copyright/imprint/warranty/trademark sections\n"
                                 "Table of contents\n"
                                 "Page numbers\n"
                                 "Headers/footers\n"
                                 "Repeated titles\n"
                                 "Broken OCR fragments\n"
                                 "Isolated headings without explanatory paragraphs\n\n"
                                 "If context quality is poor, do NOT guess.\n\n"
                                 "Say:\n"
                                 "\"The provided context does not contain enough meaningful document content to answer accurately. Please retrieve or provide relevant sections such as Introduction, Overview, Purpose, Usage, Features, Architecture, Technical Data, Connectors, Tables, or the requested component pages.\"\n\n"
                                 "If context is partially useful, answer only what is supported and clearly mark missing areas as \"Not specified in the provided context.\"\n\n"
                                 "3. Evidence and accuracy rules\n\n"
                                 "Use only information supported by the document context.\n\n"
                                 "Do not invent:\n\n"
                                 "Specifications\n"
                                 "Pin numbers\n"
                                 "Signal names\n"
                                 "Electrical values\n"
                                 "Dimensions\n"
                                 "Protocol support\n"
                                 "Features\n"
                                 "Module behavior\n"
                                 "Workflow steps\n"
                                 "Table contents\n"
                                 "Diagram details\n\n"
                                 "When you infer something from the context, label it clearly as:\n"
                                 "\"Reasonable interpretation: ...\"\n\n"
                                 "When information is absent, write:\n"
                                 "\"Not specified in the provided context.\"\n\n"
                                 "Never present guesses as facts.\n\n"
                                 "4. Noise filtering rules\n\n"
                                 "Ignore unless directly relevant:\n\n"
                                 "PDF/document metadata such as author, title, creation date\n"
                                 "Copyright, imprint, trademark, warranty, and legal notices\n"
                                 "Table of contents entries\n"
                                 "Page numbers\n"
                                 "Headers and footers\n"
                                 "Repeated section titles\n"
                                 "Raw OCR fragments\n"
                                 "Broken words caused by extraction\n"
                                 "Navigation-only text\n"
                                 "Duplicated content\n"
                                 "Lines like \"Main Features 13\", \"Important Notes 10\", or \"Page 1 Text\"\n\n"
                                 "Focus on:\n\n"
                                 "Explanatory paragraphs\n"
                                 "Product/system descriptions\n"
                                 "Purpose and intended usage\n"
                                 "Architecture and structure\n"
                                 "Functional features\n"
                                 "Components, modules, tools, or sections and their roles\n"
                                 "Workflow or operating process\n"
                                 "Inputs, outputs, interfaces, connectors, or data flow\n"
                                 "Applications and use cases\n"
                                 "Safety notes, constraints, or limitations only when meaningful\n"
                                 "Tables, diagrams, and images when the user asks for them\n\n"
                                 "5. Retrieval guidance for large documents\n\n"
                                 "For large documents, prioritize meaningful sections such as:\n\n"
                                 "Introduction\n"
                                 "Overview\n"
                                 "Purpose\n"
                                 "Intended use\n"
                                 "General description\n"
                                 "Main features\n"
                                 "Architecture\n"
                                 "System structure\n"
                                 "Usage\n"
                                 "Configuration\n"
                                 "Workflow\n"
                                 "Components/modules\n"
                                 "Connectors/interfaces\n"
                                 "Technical data\n"
                                 "Applications/use cases\n"
                                 "Safety notes\n"
                                 "Troubleshooting\n"
                                 "Requirements/specifications\n\n"
                                 "Deprioritize:\n\n"
                                 "Cover pages\n"
                                 "Imprint\n"
                                 "Copyright\n"
                                 "Warranty\n"
                                 "Trademarks\n"
                                 "Table of contents\n"
                                 "Index\n"
                                 "Repeated headers/footers\n\n"
                                 "Do not base the answer only on the first few pages unless those pages contain meaningful explanatory content.\n\n"
                                 "For component-specific requests, retrieve and use only chunks where the requested component name or aliases appear, plus nearby connector/specification/usage sections.\n\n"
                                 "For pin/diagram/table requests, prioritize pages or chunks containing:\n\n"
                                 "connector\n"
                                 "pin\n"
                                 "signal\n"
                                 "channel\n"
                                 "interface\n"
                                 "figure\n"
                                 "diagram\n"
                                 "table\n"
                                 "technical data\n"
                                 "layout\n"
                                 "wiring\n\n"
                                 "6. Format-aware handling\n\n"
                                 "Handle each document type appropriately:\n\n"
                                 "PDF:\n\n"
                                 "Preserve the meaning of figures, tables, diagrams, and page structure when available.\n"
                                 "Avoid treating table of contents as content.\n\n"
                                 "DOCX/DOC/ODT/RTF:\n\n"
                                 "Focus on headings, paragraphs, tables, and embedded images if available.\n"
                                 "Ignore repeated headers/footers.\n\n"
                                 "PPTX/PPT:\n\n"
                                 "Treat slides as structured visual content.\n"
                                 "Summarize slide intent, not just slide text.\n"
                                 "Use slide titles, bullets, diagrams, and tables together.\n\n"
                                 "XLSX/XLS/CSV:\n\n"
                                 "Identify sheets, columns, tables, metrics, and relationships.\n"
                                 "Do not summarize random cells.\n"
                                 "For analysis, explain what the data represents and key patterns if visible.\n"
                                 "For extraction, preserve rows/columns.\n\n"
                                 "HTML/Markdown/TXT:\n\n"
                                 "Use semantic headings and sections.\n"
                                 "Ignore navigation menus and boilerplate.\n\n"
                                 "Images:\n\n"
                                 "Describe visible diagrams, labels, tables, flowcharts, or screenshots.\n"
                                 "If OCR is weak, state uncertainty.\n\n"
                                 "7. Response rules by intent\n\n"
                                 "FULL_DOCUMENT_ANALYSIS\n\n"
                                 "Provide:\n\n"
                                 "Overview\n"
                                 "Purpose\n"
                                 "Core Concept\n"
                                 "Architecture / Structure\n"
                                 "Key Capabilities\n"
                                 "Major Components / Modules\n"
                                 "Workflow / How It Is Used\n"
                                 "Use Cases / Applications\n"
                                 "Important Notes / Constraints\n"
                                 "Key Takeaways\n\n"
                                 "Do not copy raw text.\n"
                                 "Do not list table-of-contents headings.\n"
                                 "Do not show page-wise extracted text.\n\n"
                                 "SHORT_SUMMARY\n\n"
                                 "Provide only:\n\n"
                                 "Short Summary\n"
                                 "What the document is about\n"
                                 "Main purpose\n"
                                 "Most important points\n"
                                 "Key takeaways\n\n"
                                 "Keep it concise.\n"
                                 "Do not include detailed architecture, long module lists, pin tables, or full feature tables unless requested.\n\n"
                                 "OVERVIEW\n\n"
                                 "Provide:\n\n"
                                 "What it is\n"
                                 "Who it is for\n"
                                 "What it is used for\n"
                                 "Main concept\n"
                                 "Main areas covered\n\n"
                                 "Keep it high-level, simple, clean, and professional.\n\n"
                                 "FEATURES_ONLY\n\n"
                                 "Extract actual functional features and capabilities.\n\n"
                                 "Do not list TOC headings such as \"Main Features 13\".\n"
                                 "Identify real features from explanatory content.\n\n"
                                 "Use:\n\n"
                                 "Feature\tWhat it does\tWhy it matters\tRelated component/module\n\n"
                                 "If a field is unavailable, write \"Not specified.\"\n\n"
                                 "SPECIFIC_COMPONENT_DETAILS\n\n"
                                 "Focus only on the requested component, module, product, section, feature, or item.\n"
                                 "Do not summarize the whole document.\n\n"
                                 "Include:\n\n"
                                 "Overview\n"
                                 "Purpose\n"
                                 "Key Features\n"
                                 "Technical Details\n"
                                 "Interfaces / Connectors\n"
                                 "Configuration / Usage\n"
                                 "Limitations / Important Notes\n"
                                 "Practical Use Cases\n"
                                 "Key Takeaways\n\n"
                                 "Ignore unrelated document content.\n\n"
                                 "PIN_DIAGRAMS_CONNECTORS_TABLES\n\n"
                                 "Focus only on visual/structural information related to the requested item.\n\n"
                                 "Include:\n\n"
                                 "Connector Overview\n"
                                 "Pin Configuration Table\n"
                                 "Channel Mapping Table, if available\n"
                                 "ASCII / structured diagram\n"
                                 "Image or figure references, if available\n"
                                 "Important notes\n\n"
                                 "Pin table format:\n\n"
                                 "Connector\tPin\tSignal / Name\tDirection\tDescription\tNotes\n\n"
                                 "Rules:\n\n"
                                 "Do not invent pin numbers, signal names, directions, or electrical values.\n"
                                 "If exact pin data is missing, clearly say: \"Exact pin data is not available in the provided context.\"\n"
                                 "Reconstruct diagrams only when the relationship is clearly supported.\n"
                                 "Make tables CSV-ready when requested.\n\n"
                                 "WORKFLOW_OR_PROCESS\n\n"
                                 "Provide:\n\n"
                                 "Process overview\n"
                                 "Step-by-step workflow\n"
                                 "Inputs\n"
                                 "Outputs\n"
                                 "Tools/components involved\n"
                                 "Practical notes\n\n"
                                 "Do not include unrelated document summary sections.\n\n"
                                 "USE_CASES_APPLICATIONS\n\n"
                                 "Provide:\n\n"
                                 "Primary use cases\n"
                                 "Real-world applications\n"
                                 "Target users\n"
                                 "Benefits\n"
                                 "Example scenarios\n\n"
                                 "COMPARISON\n\n"
                                 "Compare only the requested items.\n\n"
                                 "Use:\n\n"
                                 "Criteria\tItem 1\tItem 2\tDifference / Comment\n\n"
                                 "Include:\n\n"
                                 "Similarities\n"
                                 "Differences\n"
                                 "Best-fit usage\n"
                                 "Key takeaway\n\n"
                                 "Do not compare items that were not requested.\n\n"
                                 "TABLE_EXTRACTION\n\n"
                                 "Extract relevant tables only.\n"
                                 "Preserve rows and columns as accurately as possible.\n"
                                 "Provide Markdown table and CSV-ready format if requested.\n"
                                 "Do not summarize unless asked.\n\n"
                                 "IMAGE_OR_DIAGRAM_EXPLANATION\n\n"
                                 "Identify relevant figures, screenshots, diagrams, or visual references.\n"
                                 "Explain what each visual shows.\n"
                                 "If the image cannot be extracted, recreate a clean text-based diagram only when safe.\n\n"
                                 "DOWNLOADABLE_REPORT\n\n"
                                 "Structure content so it can be saved as:\n\n"
                                 "Markdown\n"
                                 "TXT\n"
                                 "CSV\n"
                                 "DOCX/PDF-ready report\n\n"
                                 "Clearly separate downloadable sections.\n"
                                 "Avoid decorative icons if the output is intended for CSV, TXT, DOCX, or PDF export.\n\n"
                                 "TROUBLESHOOTING_OR_LIMITATIONS\n\n"
                                 "Provide:\n\n"
                                 "Problem / limitation\n"
                                 "Likely cause from the document\n"
                                 "Relevant constraints\n"
                                 "Recommended action if stated\n"
                                 "What is not specified\n\n"
                                 "Do not invent fixes that are not supported.\n\n"
                                 "REQUIREMENTS_OR_SPECIFICATION_EXTRACTION\n\n"
                                 "Extract requirements or specifications in a structured table:\n\n"
                                 "ID\tRequirement / Specification\tCategory\tApplies to\tValue / Condition\tNotes\n\n"
                                 "Use \"Not specified\" where needed.\n\n"
                                 "8. Output style rules\n\n"
                                 "Use:\n\n"
                                 "Clean headings\n"
                                 "Professional wording\n"
                                 "Bullet points\n"
                                 "Tables where useful\n"
                                 "Concise explanations\n"
                                 "Clear separation between sections\n"
                                 "Engineering/product-documentation style\n\n"
                                 "Avoid:\n\n"
                                 "Emoji-heavy output unless the user asks\n"
                                 "Raw copied text\n"
                                 "OCR dumps\n"
                                 "Repetition\n"
                                 "Long paragraphs\n"
                                 "Unnecessary disclaimers\n"
                                 "Table of contents dumping\n"
                                 "Metadata dumping\n\n"
                                 "9. Final self-check before answering\n\n"
                                 "Before producing the final answer, verify:\n\n"
                                 "Did I answer the exact user request?\n"
                                 "Did I use the correct intent format?\n"
                                 "Did I avoid metadata, TOC, headers, footers, and OCR noise?\n"
                                 "Did I avoid copying raw text?\n"
                                 "Did I avoid repeating the same content?\n"
                                 "Did I use meaningful document content?\n"
                                 "Did I avoid inventing missing details?\n"
                                 "Did I clearly mark unsupported or missing information?\n"
                                 "Is the output professional, structured, and useful?\n"
                                 "Would \"Analyze\", \"Summary\", \"Overview\", \"Features\", \"Specific Component\", and \"Pin Diagrams\" produce clearly different outputs?\n\n"
                                 "Final instruction:\n"
                                 "Always tailor the depth, structure, and format to the user's exact query. Do not reuse the same response structure for different request types.\n\n"
                                 "INTENT CLASSIFICATION: Choose ONE primary intent only: FULL_DOCUMENT_ANALYSIS, SHORT_SUMMARY, OVERVIEW, FEATURES_ONLY, SPECIFIC_COMPONENT_DETAILS, PIN_DIAGRAMS_CONNECTORS_TABLES, WORKFLOW_OR_PROCESS, USE_CASES_APPLICATIONS, COMPARISON, TABLE_EXTRACTION, IMAGE_OR_DIAGRAM_EXPLANATION, DOWNLOADABLE_REPORT, TROUBLESHOOTING_OR_LIMITATIONS, REQUIREMENTS_OR_SPECIFICATION_EXTRACTION.\n"
                                 "   - Choose the most relevant intent. Merge only if logically necessary.\n"
                                 "   - Ambiguous priority: SPECIFIC_COMPONENT_DETAILS > COMPARISON > FULL_DOCUMENT_ANALYSIS > FEATURES_ONLY > WORKFLOW_OR_PROCESS > USE_CASES_APPLICATIONS > TABLE_EXTRACTION > IMAGE_OR_DIAGRAM_EXPLANATION > DOWNLOADABLE_REPORT > TROUBLESHOOTING_OR_LIMITATIONS > REQUIREMENTS_OR_SPECIFICATION_EXTRACTION > OVERVIEW > SHORT_SUMMARY.\n"
                                 "   - If unclear, default to SHORT_SUMMARY.\n\n"
                                 "OUTPUT FORMAT RULES:\n"
                                 "- FULL_DOCUMENT_ANALYSIS: Overview, Purpose, Core concept, Architecture / structure, Key capabilities, Major components / modules, Workflow / how it is used, Use cases / applications, Important notes / constraints, Key takeaways.\n"
                                 "- SHORT_SUMMARY: What it is, Purpose, 3–5 key insights, 2–3 key takeaways. No headings or structure references.\n"
                                 "- OVERVIEW: What it is, Who it is for, What it is used for, Main concept, Main areas covered.\n"
                                 "- FEATURES_ONLY: Table with Feature, What it does, Why it matters, Related component/module.\n"
                                 "- SPECIFIC_COMPONENT_DETAILS: Overview, Purpose, Key features, Technical details, Interfaces / connectors, Configuration / usage, Limitations / important notes, Practical use cases, Key takeaways.\n"
                                 "- PIN_DIAGRAMS_CONNECTORS_TABLES: Connector overview, Pin configuration table, Channel mapping table if available, ASCII / structured diagram, Image or figure references if available, Important notes.\n"
                                 "- WORKFLOW_OR_PROCESS: Process overview, Step-by-step workflow, Inputs, Outputs, Tools/components involved, Practical notes.\n"
                                 "- USE_CASES_APPLICATIONS: Primary use cases, Real-world applications, Target users, Benefits, Example scenarios.\n"
                                 "- COMPARISON: Comparison table first, then Similarities, Differences, Key insights, Best use cases.\n"
                                 "- TABLE_EXTRACTION: Tables only, clean schema formatting, CSV-ready if applicable, no explanation, no added fields.\n"
                                 "- IMAGE_OR_DIAGRAM_EXPLANATION: Identify relevant figures, screenshots, diagrams, or visual references. Explain what each visual shows.\n"
                                 "- DOWNLOADABLE_REPORT: Clean markdown structure, professional formatting, clearly sectioned, export-ready.\n"
                                 "- TROUBLESHOOTING_OR_LIMITATIONS: Problem / limitation, Likely cause, Relevant constraints, Recommended action, What is not specified.\n"
                                 "- REQUIREMENTS_OR_SPECIFICATION_EXTRACTION: Table with ID, Requirement / Specification, Category, Applies to, Value / Condition, Notes.\n\n"
                                 "QUALITY CONTROL: Do not repeat information across sections, do not dump raw text or OCR dumps, never invent technical values, do not mix unrelated intents. Always stay grounded in the document.\n\n"
                                 "DOCUMENT:\n{context}\n\n"
                                 "CHAT HISTORY:\n{chat_history}\n\n"
                                 "USER QUERY:\n{question}"),
                                 ("human", "{question}")
                             ])
                            chain = None
                            if llm is not None:
                                try:
                                    chain = ({"context": retriever | (lambda docs: '\n'.join(getattr(doc, "page_content", str(doc)) for doc in docs)),
                                              "chat_history": lambda _: chat_history,
                                              "document_profile": lambda _: document_profile,
                                              "question": RunnablePassthrough()} | prompt | llm)
                                except Exception as e:
                                    st.warning(f"Could not create LLM chain: {e}")
                                    chain = None

                            if chain is not None:
                                try:
                                    response = str(chain.invoke(processing_input))
                                    response = strip_llm_suggestions_from_response(response)
                                except Exception as e:
                                    st.warning(f"Could not run LLM chain: {e}")
                                    chain = None

                            if chain is None:
                                memory_hits = search_workspace_memory(processing_input, limit=4)
                                if memory_hits:
                                    response = "AI model is unavailable, so I retrieved the closest workspace memory:\n\n" + "\n\n---\n\n".join(memory_hits)
                                else:
                                    response = "⚠️ AI model is unavailable. Use direct extraction questions such as 'count(\"keyword\")', 'find(\"phrase\")', 'summarize', or 'overview'."
                        
                        # Ensure response is never empty
                        response = str(response or "").strip()
                        if not response:
                            response = "Unable to generate a response. Please try with a different query or ensure files are properly loaded."
                        if explicit_document_action and "Sources:" not in response:
                            source_query = processing_input
                            if technical_request_type in {"FULL_DOCUMENT_ANALYSIS", "SHORT_SUMMARY", "OVERVIEW"}:
                                source_query = (
                                    "introduction overview purpose main features capabilities architecture components "
                                    "workflow applications use cases technical details key takeaways"
                                )
                            citation_docs = hybrid_chatpdf_retrieve(
                                processing_input,
                                chat_files,
                                user_id=user_id,
                                final_k=7,
                                search_query=source_query,
                            )
                            sources_text = format_chatpdf_sources(citation_docs) if citation_docs else "\n".join(
                                f"- {file_name}" for file_name in chat_files
                            )
                            response = response.rstrip() + "\n\nSources:\n" + sources_text
                        
                        current_chat_messages.append({"role": "assistant", "content": response})
                        st.session_state.document_chat_display[chat_display_key] = current_chat_messages[-100:]
                        st.session_state.messages = st.session_state.document_chat_display[chat_display_key]
                        st.session_state.last_streamed_assistant_index = len(st.session_state.messages) - 1
                        st.session_state.chat_next_suggestions = build_chat_next_suggestions(processing_input, combined_text, chat_intent)
                        st.session_state.chat_next_suggestions_for = len(st.session_state.messages) - 1
                        append_chat_to_workspace_memory(submitted_input, response, chat_files)
                        save_workspace_memory()
                        save_memory_log("chat", "Chat interaction stored in workspace memory.", {
                            "user": submitted_input,
                            "files": chat_files,
                            "assistant_preview": response[:300],
                        })
                        if "⚠️" in response or "not found" in response.lower() or "please select" in response.lower() or "ai model is unavailable" in response.lower():
                            set_help_popup_state("chat", True)

        visible_messages = locals().get("current_chat_messages", st.session_state.get("messages", []))
        for msg_index, raw_msg in enumerate(visible_messages):
            msg = get_valid_chat_message(raw_msg)
            if msg is None:
                continue

            avatar = "🧑" if msg["role"] == "user" else "🤖"
            with st.chat_message(msg["role"], avatar=avatar):
                render_chat_message_content(msg, msg_index)

            if (
                msg["role"] == "assistant"
                and msg_index == st.session_state.get("chat_next_suggestions_for")
                and st.session_state.get("chat_next_suggestions")
            ):
                st.caption("Suggested next steps")
                next_suggestions = list(dict.fromkeys(st.session_state.get("chat_next_suggestions", [])))[:4]
                # Ensure all suggestions are non-empty strings
                next_suggestions = [str(s).strip() for s in next_suggestions if s and str(s).strip()]
                
                if next_suggestions:
                    suggestion_cols = st.columns(len(next_suggestions))
                    for suggestion_index, suggestion_text in enumerate(next_suggestions):
                        if suggestion_text:  # Double-check suggestion is not empty
                            with suggestion_cols[suggestion_index]:
                                if st.button(
                                    suggestion_text,
                                    key=f"ai_sugg_{msg_index}_{suggestion_index}",
                                    use_container_width=True,
                                ):
                                    st.session_state.input_prefill = suggestion_text
                                    st.rerun()

        render_chat_summary_downloads()
    else:
        st.info("Select files from the sidebar to start chatting.")

    st.markdown('</div>', unsafe_allow_html=True)

    # -------------------------------
