# Auto-generated from legacy_app.py during modular refactor.
# The original monolith is retained as rollback documentation.

from functions import *

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

    def build_chat_next_suggestions(user_input, context):
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
            if "table" in memory_text or "data" in memory_text:
                suggestions.append("Inspect table data")

        entity = extract_chat_entity(user_input)
        if entity:
            if any(term in user_input.lower() for term in ["pin", "diagram", "connector", "d-sub"]):
                suggestions.insert(0, f"Pin diagram: {entity}")
            else:
                suggestions.insert(0, f"Item details: {entity}")

        if not suggestions:
            suggestions = [
                "Get document overview",
                "List key facts",
                "Find specific terms",
                "Explain structure",
            ]
        return list(dict.fromkeys(suggestions))[:4]

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
    show_current_sidebar_selection()
    render_file_context_card("Chat File Context", st.session_state.selected_files, st.session_state.chat_file_selection)

    show_help_popup('chat', st.session_state.selected_files)

    if st.session_state.selected_files:
        st.session_state.chat_file_selection = [
            file_name for file_name in st.session_state.chat_file_selection
            if file_name in st.session_state.selected_files
        ]
        chat_files = st.multiselect("Choose file(s) for Chat", options=st.session_state.selected_files,
                                    default=st.session_state.chat_file_selection, key="chat_file_selection")
        if not chat_files:
            st.info("Choose one or more files in this tab to start chatting.")
        else:
            with st.spinner("Loading selected files..."):
                ensure_files_processed(chat_files)
            combined_text = "\n".join([st.session_state.file_texts.get(f, "") for f in chat_files])


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
                    st.session_state.messages = []
                    st.session_state.chat_summary_downloads = empty_chat_summary_downloads()
                    st.session_state.chat_next_suggestions = []
                    st.session_state.chat_next_suggestions_for = None
                    st.success("✅ Chat cleared!")
                else:
                    st.session_state.messages.append({"role": "user", "content": submitted_input})
                    with st.spinner("Processing your request..."):
                        st.session_state.chat_summary_downloads = empty_chat_summary_downloads()
                        user_input_lower = processing_input.lower()
                        # Word count queries
                        if any(t in user_input_lower for t in ["how many", "count", "number of", "occurrences"]):
                            match = re.search(r"'(.*?)'|\"(.*?)\"", processing_input)
                            if match:
                                word = match.group(1) or match.group(2)
                                count = len(
                                    re.findall(rf'(?<![\w-]){re.escape(word)}(?![\w-])', combined_text, re.IGNORECASE))
                                response = f"🔢 The word/phrase '{word}' appears {count} times in the selected documents."
                            else:
                                response = "⚠️ Specify the word/phrase in quotes. Example: count('keyword') or count(\"keyword\")"
                        elif any(term in user_input_lower for term in ["find", "search", "locate"]) or "highlight" in user_input_lower:
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
                        elif "overview" in user_input_lower:
                            response_lines = []
                            for f in chat_files:
                                file_text = st.session_state.file_texts.get(f, "")
                                if file_text.strip():
                                    toc_entries = extract_toc_with_page_numbers(file_text)
                                    all_headings = extract_document_headings(file_text)
                                    if all_headings:
                                        response_lines.append(f"📄 **{f}**")
                                        response_lines.append("### Document Headings")
                                        for num, title in all_headings:
                                            content_str = f"{num} {title}" if num else title
                                            page_num = resolve_heading_page_number(file_text, title, toc_entries)
                                            display_text = f"{content_str} (Page {page_num})" if page_num else content_str
                                            preview_link = create_preview_link(f, highlight_term=title, page_num=page_num)
                                            if preview_link:
                                                anchor_id = create_heading_anchor(title)
                                                response_lines.append(f"- <a href='{preview_link}#{anchor_id}' target='_blank'>{html.escape(display_text)}</a>")
                                            else:
                                                response_lines.append(f"- {html.escape(display_text)}")
                                    else:
                                        response_lines.append(f"📄 **{f}**\n\nNo document headings were detected.")
                                else:
                                    response_lines.append(f"📄 **{f}**\n\nNo readable content found in this document.")
                            response = "\n\n".join(response_lines)
                        elif any(term in user_input_lower for term in ["pin diagram", "pin table", "pin configuration", "visual reference", "visual and structural", "connector details", "technical tables"]):
                            item_name = extract_quoted_item_name(processing_input)
                            if not item_name:
                                response = "⚠️ Specify the item name in quotes. Example: pin diagram \"D-SUB9\" or visual reference \"VN1640A\""
                            else:
                                response_blocks = []
                                pin_csv_downloads = []
                                ascii_diagram_downloads = []
                                for f in chat_files:
                                    file_text = st.session_state.file_texts.get(f, "")
                                    if file_text.strip():
                                        response_blocks.append(build_item_visual_response(f, file_text, item_name))
                                        visual_assets = build_item_visual_assets(f, file_text, item_name)
                                        pin_csv_downloads.extend(visual_assets.get("csv", []))
                                        ascii_diagram_downloads.extend(visual_assets.get("diagrams", []))
                                    else:
                                        response_blocks.append(f"📄 **{f}**\n\nNo readable content found in this document.")
                                st.session_state.chat_summary_downloads = {
                                    "images": [],
                                    "tables": [],
                                    "csv": pin_csv_downloads,
                                    "diagrams": ascii_diagram_downloads,
                                }
                                response = "\n\n---\n\n".join(response_blocks)
                        elif any(term in user_input_lower for term in ["item details", "item information", "extract item", "about item", "information about", "details about"]):
                            item_name = extract_quoted_item_name(processing_input)
                            if not item_name:
                                response = "⚠️ Specify the item name in quotes. Example: item details \"VN1630A\" or information about \"D-SUB9\""
                            else:
                                response_blocks = []
                                for f in chat_files:
                                    file_text = st.session_state.file_texts.get(f, "")
                                    if file_text.strip():
                                        response_blocks.append(build_item_information_response(f, file_text, item_name))
                                    else:
                                        response_blocks.append(f"📄 **{f}**\n\nNo readable content found in this document.")
                                response = "\n\n---\n\n".join(response_blocks)
                        elif extract_bare_item_name(processing_input):
                            item_name = extract_bare_item_name(processing_input)
                            response_blocks = []
                            for f in chat_files:
                                file_text = st.session_state.file_texts.get(f, "")
                                if file_text.strip():
                                    response_blocks.append(build_item_information_response(f, file_text, item_name))
                                else:
                                    response_blocks.append(f"ðŸ“„ **{f}**\n\nNo readable content found in this document.")
                            response = "\n\n---\n\n".join(response_blocks)
                        elif any(term in user_input_lower for term in ["analyze", "summary", "summarize", "summarise"]):
                            result = []
                            summary_image_downloads = []
                            summary_table_downloads = []
                            for f in chat_files:
                                file_text = st.session_state.file_texts.get(f, "")
                                file_entry = get_uploaded_file_entry(f)
                                if file_text.strip():
                                    file_bytes = file_entry["bytes"] if file_entry else b""
                                    result.append(build_detailed_document_summary(f, file_bytes, file_text))
                                    page_match = re.search(r"Total Pages:\s*(\d+)", file_text)
                                    page_count = int(page_match.group(1)) if page_match else 0
                                    is_large_pdf = f.lower().endswith(".pdf") and page_count > PDF_ASSET_SCAN_PAGE_LIMIT
                                    if not is_large_pdf:
                                        summary_assets = build_summary_download_assets(f, file_bytes)
                                        summary_image_downloads.extend(summary_assets.get("images", []))
                                        summary_table_downloads.extend(summary_assets.get("tables", []))
                                else:
                                    result.append(f"📄 **{f}**\n\nNo readable content found in this document.")

                            st.session_state.chat_summary_downloads = {
                                "images": summary_image_downloads,
                                "tables": summary_table_downloads,
                                "csv": [],
                                "diagrams": [],
                            }
                            response = "\n\n---\n\n".join(result)
                        elif "compare" in user_input_lower:
                            if len(chat_files) >= 2:
                                selected_texts = {f: st.session_state.file_texts[f] for f in chat_files}
                                response = highlight_multi_file_differences(selected_texts)
                            else:
                                response = "⚠️ Please select at least 2 files to compare documents."
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
                                 "You are a ChatGPT-like AI assistant specialized in analyzing user-uploaded documents.\n\nYour job is to:\n1. Understand the document\n2. Respond conversationally to the user\n3. Generate smart, context-aware suggestions\n\n---\n\nSTEP 1: DOCUMENT UNDERSTANDING\n\nIdentify:\n- Document type (e.g., dataset, invoice, technical doc, legal, general text)\n- Structure (structured, semi-structured, unstructured)\n- Key elements (tables, IDs, components, signals, dates, totals, etc.)\n\n---\n\nSTEP 2: USER INTENT\n\n- Understand the user query\n- If no query is provided, assume the user wants an overview\n\n---\n\nSTEP 3: RESPONSE (ChatGPT-style)\n\n- Respond naturally and clearly\n- Be concise but informative\n- Provide insights, not just surface-level answers\n- If data is present, mention trends or patterns\n- If unsure, say what is missing\n\n---\n\nSTEP 4: SUGGESTIONS\n\nGenerate 4–6 highly relevant next actions.\nRules:\n- Each suggestion must be short (2–5 words)\n- Must be specific to the document\n- Must reflect real user actions\n- Avoid generic phrases like \"Analyze document\"\n\nExamples:\nDataset:\n- Analyze trends\n- Find anomalies\n- Show summary stats\n\nInvoice:\n- Extract total\n- Find due date\n- List line items\n\nTechnical:\n- Explain components\n- Show pin diagram\n- List signals\n\n---\n\nSTEP 5: ENTITY AWARENESS (IMPORTANT)\n\nIf the document or user query contains a clear entity (e.g., part number, item ID, component name):\n- Adapt suggestions using that entity\n\nExample:\nInstead of:\n- Show item details\nUse:\n- Item details: VN1630A\n\nInstead of:\n- Show pin diagram\nUse:\n- Pin diagram: D-SUB9\n\n---\n\nOUTPUT FORMAT:\nReturn your response in this structure:\n\n<chat_response>\n\n---\nSuggestions:\n- suggestion 1\n- suggestion 2\n- suggestion 3\n- suggestion 4\n\n---\n\nDOCUMENT:\n{context}\n\nCONVERSATION HISTORY:\n{chat_history}\n\nUSER MESSAGE:\n{question}"),
                                ("human", "{question}")
                            ])
                            chain = None
                            if llm is not None:
                                try:
                                    chain = ({"context": retriever | (lambda docs: '\n'.join(getattr(doc, "page_content", str(doc)) for doc in docs)),
                                              "chat_history": lambda _: chat_history,
                                              "question": RunnablePassthrough()} | prompt | llm)
                                except Exception as e:
                                    st.warning(f"Could not create LLM chain: {e}")
                                    chain = None

                            if chain is not None:
                                try:
                                    response = str(chain.invoke(processing_input))
                                except Exception as e:
                                    st.warning(f"Could not run LLM chain: {e}")
                                    chain = None

                            if chain is None:
                                memory_hits = search_workspace_memory(processing_input, limit=4)
                                if memory_hits:
                                    response = "AI model is unavailable, so I retrieved the closest workspace memory:\n\n" + "\n\n---\n\n".join(memory_hits)
                                else:
                                    response = "⚠️ AI model is unavailable. Use direct extraction questions such as 'count(\"keyword\")', 'find(\"phrase\")', 'summarize', or 'overview'."
                        st.session_state.messages.append({"role": "assistant", "content": response})
                        st.session_state.last_streamed_assistant_index = len(st.session_state.messages) - 1
                        st.session_state.chat_next_suggestions = build_chat_next_suggestions(processing_input, combined_text)
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

        for msg_index, msg in enumerate(st.session_state.messages):
            role = "🧑" if msg["role"] == "user" else "🤖"
            st.markdown(f"**{role}**", unsafe_allow_html=True)
            if msg["role"] == "assistant" and msg_index == st.session_state.get("last_streamed_assistant_index"):
                placeholder = st.empty()
                content = str(msg["content"])
                tokens = re.split(r"(\s+)", content)
                streamed = ""
                for token_index, token in enumerate(tokens):
                    streamed += token
                    if token_index < 240:
                        placeholder.markdown(streamed + "▌", unsafe_allow_html=True)
                        time.sleep(0.006)
                placeholder.markdown(content, unsafe_allow_html=True)
                st.session_state.last_streamed_assistant_index = None
            else:
                st.markdown(msg["content"], unsafe_allow_html=True)

            if (
                msg["role"] == "assistant"
                and msg_index == st.session_state.get("chat_next_suggestions_for")
                and st.session_state.get("chat_next_suggestions")
            ):
                st.caption("Suggested next steps")
                next_suggestions = list(dict.fromkeys(st.session_state.get("chat_next_suggestions", [])))[:4]
                suggestion_cols = st.columns(len(next_suggestions))
                for suggestion_index, suggestion_text in enumerate(next_suggestions):
                    if suggestion_cols[suggestion_index].button(
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
