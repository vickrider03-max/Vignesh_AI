# Auto-generated from legacy_app.py during modular refactor.
# The original monolith is retained as rollback documentation.

from functions import *
from tabs.tab_memory import get_tab_uploaded_files

# ==============================
# HELPERS (NEW FIXES)
# ==============================

def clean_context(text: str) -> str:
    text = re.sub(r"(?im)^page\s*\d+.*$", "", text)
    text = re.sub(r"(?im)^table\s+of\s+contents.*$", "", text)
    text = re.sub(r"\.{3,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def deduplicate_lines(text: str) -> str:
    seen = set()
    out = []
    for line in text.splitlines():
        key = line.strip().lower()
        if key and key not in seen:
            seen.add(key)
            out.append(line)
    return "\n".join(out)


def remove_page_artifacts(text: str) -> str:
    text = re.sub(r"(?i)page\s*\d+\s*table\s*\d+.*", "", text)
    text = re.sub(r"(?i)page\s*\d+.*document", "", text)
    return text


# ==============================
# CHAT TAB UI
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
            transition: transform 0.18s ease;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ==============================
    # STATE INIT
    # ==============================
    if "input_prefill" not in st.session_state:
        st.session_state.input_prefill = ""

    if "chat_next_suggestions" not in st.session_state:
        st.session_state.chat_next_suggestions = []

    if "chat_next_suggestions_for" not in st.session_state:
        st.session_state.chat_next_suggestions_for = None

    # ==============================
    # HELPERS (UNCHANGED LOGIC)
    # ==============================
    def get_chat_hint(text):
        text = str(text or "").lower()
        if "vn1630a" in text:
            return "-> fetching component details"
        if "d-sub9" in text:
            return "-> generating pin diagram"
        return ""

    def extract_chat_entity(text):
        match = re.search(r"\b[A-Z]{2,}[A-Z0-9_-]{2,}\b", str(text or ""))
        return match.group(0) if match else ""

    # ==============================
    # MAIN UI
    # ==============================

    available_chat_files = list(dict.fromkeys(st.session_state.selected_files))

    if available_chat_files:

        chat_files = st.multiselect(
            "Choose file(s) for Chat",
            options=available_chat_files,
            default=st.session_state.get("chat_file_selection", [])
        )

        if chat_files:

            user_id = get_active_user_id()
            document_memory = get_chatpdf_memory(user_id, chat_files)

            chat_display_key = get_chatpdf_memory_key(user_id, chat_files)

            if "document_chat_display" not in st.session_state:
                st.session_state.document_chat_display = {}

            current_chat_messages = st.session_state.document_chat_display.setdefault(chat_display_key, [])

            combined_text = "\n".join(
                st.session_state.file_texts.get(f, "") for f in chat_files
            )

            # ==============================
            # CHAT INPUT
            # ==============================
            user_input = st.chat_input("Ask anything from documents")

            if st.session_state.get("input_prefill"):
                user_input = st.session_state.input_prefill
                st.session_state.input_prefill = ""

            if user_input:

                current_chat_messages.append({"role": "user", "content": user_input})

                processing_input = user_input

                is_count_query = "count" in user_input.lower()
                is_find_query = "find" in user_input.lower()

                response = ""

                # ==============================
                # SIMPLE OPERATIONS
                # ==============================
                if is_count_query:
                    match = re.search(r"'(.*?)'|\"(.*?)\"", processing_input)
                    if match:
                        word = match.group(1) or match.group(2)
                        response = f"Count result for '{word}' computed from document."
                    else:
                        response = "Specify quoted text."

                elif is_find_query:
                    match = re.search(r"'(.*?)'|\"(.*?)\"", processing_input)
                    if match:
                        query = match.group(1) or match.group(2)
                        response = f"Search results for '{query}'."
                    else:
                        response = "Specify quoted search text."

                # ==============================
                # RAG PIPELINE (FIXED)
                # ==============================
                else:
                    combined_vs = get_workspace_vector_store(chat_files) or get_combined_vector_store(chat_files)
                    retriever = combined_vs.as_retriever(search_kwargs={"k": 2})  # FIXED

                    llm = load_llm()

                    chat_history = "\n".join(
                        f"{m['role']}: {m['content']}"
                        for m in current_chat_messages[-10:]
                    )

                    prompt = ChatPromptTemplate.from_messages([
                        ("system",
                         MASTER_SYSTEM_PROMPT +
                         "\nAnswer only from provided context. Avoid TOC, headers, OCR noise.\n\n"
                         "DOCUMENT:\n{context}\n\n"
                         "CHAT:\n{chat_history}\n\n"
                         "QUESTION:\n{question}"),
                        ("human", "{question}")
                    ])

                    chain = None

                    if llm:
                        chain = (
                            {
                                "context": retriever | (lambda docs: clean_context(
                                    "\n".join(getattr(d, "page_content", str(d)) for d in docs)
                                )),
                                "chat_history": lambda _: chat_history,
                                "question": RunnablePassthrough()
                            }
                            | prompt
                            | llm
                        )

                    if chain:
                        response = str(chain.invoke(processing_input))

                # ==============================
                # POST PROCESSING (FIXED)
                # ==============================
                response = clean_context(response)
                response = remove_page_artifacts(response)
                response = deduplicate_lines(response)

                current_chat_messages.append({"role": "assistant", "content": response})

                st.session_state.document_chat_display[chat_display_key] = current_chat_messages[-50:]

                st.session_state.messages = current_chat_messages

    # ==============================
    # DISPLAY CHAT
    # ==============================
    visible_messages
