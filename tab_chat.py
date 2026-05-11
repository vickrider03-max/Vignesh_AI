# New document-chat tab implementation.
# It reuses the existing extraction, semantic file-brain, and grounded retrieval
# helpers from functions.py while keeping the Chat tab UI focused and predictable.

from functions import *


SUPPORTED_CHAT_FORMATS = {
    "pdf",
    "doc",
    "docx",
    "txt",
    "md",
    "log",
    "ppt",
    "pptx",
    "xls",
    "xlsx",
    "csv",
    "html",
    "htm",
    "odt",
    "rtf",
    "pages",
    "capl",
    "can",
    "png",
    "jpg",
    "jpeg",
    "gif",
    "bmp",
    "webp",
}


ADVANCED_DOCUMENT_CHAT_RULES = """
You are an advanced AI document analysis and conversational retrieval system.

Core behavior:
- Understand uploaded documents deeply.
- Answer only from selected document knowledge.
- Preserve technical correctness.
- Say when the document does not provide the requested information.
- Maintain conversational continuity.
- Connect related sections when the document supports the connection.
- Adapt to technical, business, legal, financial, and research documents.
"""


def _file_extension(file_name):
    return os.path.splitext(str(file_name or "").strip().lower())[1].lstrip(".")


def _is_supported_chat_file(file_name):
    return _file_extension(file_name) in SUPPORTED_CHAT_FORMATS


def _display_file_type(file_name):
    ext = _file_extension(file_name)
    family = detect_file_type(file_name)
    if ext:
        return f"{ext.upper()} ({family})"
    return family.title()


def _domain_from_semantic(semantic):
    domains = []
    for item in semantic.get("technical_domains", []) or []:
        if isinstance(item, dict) and item.get("domain"):
            domains.append(str(item.get("domain")))
        elif item:
            domains.append(str(item))
    return ", ".join(dict.fromkeys(domains)[:3]) if domains else "General document"


def _technical_level_from_text(text, semantic):
    combined = " ".join(
        [
            str(text or "")[:9000],
            " ".join(str(item) for item in semantic.get("key_concepts", [])[:20]),
            " ".join(str(item) for item in semantic.get("architecture_components", [])[:20]),
        ]
    ).lower()
    if any(term in combined for term in ["protocol", "architecture", "interface", "diagnostic", "api", "capl", "connector", "signal"]):
        return "Technical"
    if any(term in combined for term in ["methodology", "experiment", "dataset", "findings", "hypothesis", "evaluation"]):
        return "Research"
    if any(term in combined for term in ["revenue", "risk", "strategy", "contract", "compliance", "financial", "legal"]):
        return "Business / analytical"
    return "General"


def _purpose_from_semantic(semantic):
    summary = normalize_synthesis_text(
        semantic.get("executive_summary")
        or semantic.get("document_summary")
        or semantic.get("technical_summary")
        or ""
    )
    if not summary:
        return "Purpose not explicit in the readable content."
    first_sentence = re.split(r"(?<=[.!?])\s+", summary.strip())[0]
    return first_sentence[:260]


def _audience_from_level(level, semantic):
    concepts = " ".join(str(item) for item in semantic.get("key_concepts", [])[:20]).lower()
    if level == "Technical":
        if any(term in concepts for term in ["capl", "can", "diagnostic", "signal", "interface", "connector"]):
            return "Engineers, developers, testers, or technical reviewers"
        return "Technical and implementation teams"
    if level == "Research":
        return "Researchers, analysts, and subject-matter reviewers"
    if level == "Business / analytical":
        return "Business stakeholders, analysts, reviewers, or decision makers"
    return "General readers and document users"


def _build_document_profiles(file_names, file_brains):
    profiles = []
    for file_name in file_names or []:
        brain = (file_brains or {}).get(file_name, {}) or {}
        semantic = brain.get("semantic_metadata", {}) or {}
        extracted_text = st.session_state.get("file_texts", {}).get(file_name, "")
        level = _technical_level_from_text(extracted_text, semantic)
        profiles.append(
            {
                "name": file_name,
                "extension": _file_extension(file_name) or "unknown",
                "type": _display_file_type(file_name),
                "domain": _domain_from_semantic(semantic),
                "purpose": _purpose_from_semantic(semantic),
                "audience": _audience_from_level(level, semantic),
                "technical_level": level,
                "concepts": [str(item) for item in semantic.get("key_concepts", [])[:6] if str(item).strip()],
            }
        )
    return profiles


def _render_profile_cards(profiles):
    if not profiles:
        return

    cards = []
    for profile in profiles:
        concept_html = "".join(
            f"<span class='chat-profile-chip'>{html.escape(concept)}</span>"
            for concept in profile["concepts"]
        )
        cards.append(
            f"""
            <div class="chat-profile-card">
                <div class="chat-profile-title">{html.escape(profile["name"])}</div>
                <div class="chat-profile-meta">
                    <span>{html.escape(profile["type"])}</span>
                    <span>{html.escape(profile["domain"])}</span>
                    <span>{html.escape(profile["technical_level"])}</span>
                </div>
                <p><b>Purpose:</b> {html.escape(profile["purpose"])}</p>
                <p><b>Audience:</b> {html.escape(profile["audience"])}</p>
                <div class="chat-profile-chip-wrap">
                    {concept_html if concept_html else "<span class='chat-profile-chip'>Concepts not detected yet</span>"}
                </div>
            </div>
            """
        )

    st.markdown("".join(cards), unsafe_allow_html=True)


def _chat_key_for_files(user_id, file_names):
    try:
        return get_chatpdf_memory_key(user_id, file_names)
    except Exception:
        signature = "::".join(sorted(str(name) for name in file_names or []))
        return f"advanced_chat::{user_id}::{signature}"


def _normalize_chat_response(response_text):
    text = str(response_text or "").strip()
    if not text:
        return MISSING_DOCUMENT_INFO_MESSAGE
    text = strip_llm_suggestions_from_response(text)
    text = re.sub(r"(?i)\bnot specified in the provided context\.?", MISSING_DOCUMENT_INFO_MESSAGE, text)
    text = re.sub(r"(?i)\bi cannot answer from the retrieved context\.?", MISSING_DOCUMENT_INFO_MESSAGE, text)
    return text.strip()


def _answer_selected_documents(question, chat_files, user_id):
    """Ground every answer in the selected uploaded files."""
    response, citation_docs = answer_chatpdf_question(
        str(question or "").strip(),
        chat_files,
        user_id=user_id,
        top_k=10,
    )
    return _normalize_chat_response(response), citation_docs


def _render_chat_css():
    st.markdown(
        """
        <style>
        .advanced-chat-shell {
            border: 1px solid rgba(15, 23, 42, 0.10);
            background: #ffffff;
            border-radius: 8px;
            padding: 14px 16px;
            margin: 0 0 14px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
        }
        .advanced-chat-title {
            font-size: 1.08rem;
            font-weight: 800;
            color: #0f172a;
            margin: 0 0 4px;
        }
        .advanced-chat-subtitle {
            color: #475569;
            font-size: 0.92rem;
            margin: 0;
            line-height: 1.45;
        }
        .chat-profile-card {
            border: 1px solid rgba(148, 163, 184, 0.35);
            border-radius: 8px;
            background: #f8fbff;
            padding: 12px 14px;
            margin: 10px 0;
        }
        .chat-profile-title {
            font-weight: 800;
            color: #173152;
            overflow-wrap: anywhere;
            margin-bottom: 6px;
        }
        .chat-profile-meta {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
            margin-bottom: 8px;
        }
        .chat-profile-meta span,
        .chat-profile-chip {
            display: inline-flex;
            align-items: center;
            border: 1px solid #cfe2f3;
            background: #ffffff;
            color: #173152;
            border-radius: 999px;
            padding: 0.22rem 0.55rem;
            font-size: 0.78rem;
            font-weight: 700;
        }
        .chat-profile-card p {
            color: #334155;
            font-size: 0.9rem;
            line-height: 1.45;
            margin: 0.35rem 0;
        }
        .chat-profile-chip-wrap {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
            margin-top: 8px;
        }
        .st-key-chat_file_selection div[data-baseweb="select"] > div {
            background: #ffffff !important;
            border: 1px solid #cfe2f3 !important;
            border-radius: 8px !important;
            min-height: 44px !important;
        }
        [data-testid="stChatInput"] > div {
            border-radius: 12px !important;
            border: 1px solid #cfe2f3 !important;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_chat_tab():
    """Render the advanced document chat tab."""
    _render_chat_css()

    for key, default in {
        "chat_file_selection": [],
        "document_chat_display": {},
        "chat_next_suggestions": [],
        "chat_summary_downloads": empty_chat_summary_downloads(),
    }.items():
        if key not in st.session_state:
            st.session_state[key] = default

    header_col, reset_col = st.columns([8, 1])
    with header_col:
        st.markdown(
            """
            <div class="advanced-chat-shell">
                <div class="advanced-chat-title">Chat with Selected Documents</div>
                <p class="advanced-chat-subtitle">
                    Select uploaded files, review their detected document type, then ask grounded questions.
                    Answers are generated only from the selected document context.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with reset_col:
        if st.button("Reset", key="reset_chat_selection", help="Clear selected chat files and conversation"):
            st.session_state.chat_file_selection = []
            st.session_state.document_chat_display = {}
            st.session_state.messages = []
            st.session_state.chat_next_suggestions = []
            st.session_state.chat_summary_downloads = empty_chat_summary_downloads()
            st.rerun()

    sidebar_files = list(dict.fromkeys(st.session_state.get("selected_files", [])))
    supported_sidebar_files = [file_name for file_name in sidebar_files if _is_supported_chat_file(file_name)]
    unsupported_sidebar_files = [file_name for file_name in sidebar_files if not _is_supported_chat_file(file_name)]

    show_current_sidebar_selection()
    if unsupported_sidebar_files:
        st.warning("These selected files are not supported by Chat: " + ", ".join(unsupported_sidebar_files))

    st.session_state.chat_file_selection = [
        file_name for file_name in st.session_state.get("chat_file_selection", [])
        if file_name in supported_sidebar_files
    ]
    render_file_context_card("Chat File Context", supported_sidebar_files, st.session_state.chat_file_selection)

    if not supported_sidebar_files:
        st.info("Upload files in the sidebar and click their file cards. The selected files will appear here.")
        st.caption("Supported formats: " + ", ".join(sorted(SUPPORTED_CHAT_FORMATS)))
        return

    chat_files = st.multiselect(
        "Choose file(s) for Chat",
        options=supported_sidebar_files,
        default=st.session_state.chat_file_selection,
        key="chat_file_selection",
        help="Only files selected from the sidebar Uploaded files list are available here.",
    )

    if not chat_files:
        st.info("Choose one or more files to start a grounded document chat.")
        return

    with st.spinner("Identifying document types and building document understanding..."):
        ensure_files_processed(chat_files)
        file_brains = get_file_brains(chat_files)

    profiles = _build_document_profiles(chat_files, file_brains)
    with st.expander("Detected document understanding", expanded=True):
        _render_profile_cards(profiles)
        render_document_understanding_panel(file_brains)

    user_id = get_active_user_id()
    chat_display_key = _chat_key_for_files(user_id, chat_files)
    current_messages = st.session_state.document_chat_display.setdefault(chat_display_key, [])
    st.session_state.messages = current_messages

    for index, message in enumerate(current_messages):
        role = str(message.get("role", "")).strip().lower()
        content = str(message.get("content", "") or "").strip()
        if role not in {"user", "assistant"} or not content:
            continue
        with st.chat_message(role):
            if role == "assistant":
                st.markdown(content, unsafe_allow_html=True)
            else:
                st.markdown(content)

    prompt = st.chat_input("Ask a question about the selected document(s)")
    if not prompt:
        st.caption("Ask about meaning, architecture, workflow, risks, comparisons, insights, or exact details.")
        return

    current_messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Reading the selected document context..."):
            response, citation_docs = _answer_selected_documents(prompt, chat_files, user_id)
        st.markdown(response, unsafe_allow_html=True)

    current_messages.append({"role": "assistant", "content": response})
    st.session_state.document_chat_display[chat_display_key] = current_messages[-100:]
    st.session_state.messages = st.session_state.document_chat_display[chat_display_key]
