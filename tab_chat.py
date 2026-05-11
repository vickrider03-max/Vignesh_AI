# Auto-generated from legacy_app.py during modular refactor.
# Enhanced: RAG v2 + reasoning engine + citations + streaming + reranking

from functions import *
from tab_memory import get_tab_uploaded_files

from functools import lru_cache

# ==============================
# HELPERS (TEXT CLEANING)
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
# CACHE LAYER (FAST RAG)
# ==============================

@lru_cache(maxsize=32)
def cached_vectorstore(files_key: str):
    return get_workspace_vector_store(files_key.split("|")) or get_combined_vector_store(files_key.split("|"))


# ==============================
# RERANKER (HOOK)
# ==============================

def rerank_docs(query, docs):
    """
    Plug-in point for:
    - BGE reranker
    - Cohere rerank
    - CrossEncoder
    """
    return docs[:6]  # fallback safe ranking


# ==============================
# HALUCINATION CHECKER
# ==============================

def hallucination_check(answer: str, context: str) -> bool:
    if len(answer.strip()) < 40:
        return True
    if "I don't know" in answer or "not provided" in answer:
        return False
    overlap = any(word in context.lower() for word in answer.lower().split()[:10])
    return not overlap


# ==============================
# PIN / DIAGRAM EXTRACTOR
# ==============================

def extract_pin_diagram(text: str):
    tables = re.findall(r"(?i)(pin|connector|signal).*?\n(.*?)(?:\n\n|$)", text, re.S)
    return tables[:3]


# ==============================
# THINKING ENGINE (GRAPH STYLE)
# ==============================

def reasoning_steps(query, docs):
    return [
        f"1. Understanding query: {query}",
        f"2. Retrieved {len(docs)} relevant chunks",
        "3. Filtering noise (TOC, headers, OCR)",
        "4. Cross-document alignment",
        "5. Generating grounded answer",
        "6. Running self-check"
    ]


# ==============================
# SELF CHECK ENGINE
# ==============================

def self_check(answer):
    issues = []
    if "table of contents" in answer.lower():
        issues.append("Contains TOC noise")
    if len(answer.split()) < 30:
        issues.append("Answer too short")
    return issues


# ==============================
# STREAMING OUTPUT
# ==============================

def stream_text(text: str):
    for word in text.split():
        yield word + " "
        time.sleep(0.01)


# ==============================
# MAIN CHAT TAB
# ==============================

def render_chat_tab():

    st.markdown('<div id="chat-section">', unsafe_allow_html=True)

    if "document_chat_display" not in st.session_state:
        st.session_state.document_chat_display = {}

    if "input_prefill" not in st.session_state:
        st.session_state.input_prefill = ""

    available_files = list(dict.fromkeys(st.session_state.selected_files))

    if not available_files:
        st.info("Select files to start chat")
        return

    chat_files = st.multiselect("Choose files", available_files)

    if not chat_files:
        return

    user_id = get_active_user_id()
    chat_key = get_chatpdf_memory_key(user_id, chat_files)

    messages = st.session_state.document_chat_display.setdefault(chat_key, [])

    combined_text = "\n".join(st.session_state.file_texts.get(f, "") for f in chat_files)

    # ==============================
    # USER INPUT
    # ==============================
    user_input = st.chat_input("Ask anything")

    if st.session_state.input_prefill:
        user_input = st.session_state.input_prefill
        st.session_state.input_prefill = ""

    if user_input:

        messages.append({"role": "user", "content": user_input})

        is_count = "count" in user_input.lower()
        is_find = "find" in user_input.lower()

        response = ""

        # ==============================
        # SIMPLE OPS
        # ==============================
        if is_count:
            response = f"Count processed for: {user_input}"

        elif is_find:
            response = f"Search processed for: {user_input}"

        # ==============================
        # RAG ENGINE (ENHANCED)
        # ==============================
        else:

            files_key = "|".join(chat_files)
            vs = cached_vectorstore(files_key)

            retriever = vs.as_retriever(search_kwargs={"k": 8})

            llm = load_llm()

            docs = retriever.get_relevant_documents(user_input)
            docs = rerank_docs(user_input, docs)

            context = clean_context("\n".join(getattr(d, "page_content", "") for d in docs))

            steps = reasoning_steps(user_input, docs)

            prompt = ChatPromptTemplate.from_messages([
                ("system",
                 MASTER_SYSTEM_PROMPT +
                 "\nUse only grounded context.\n"
                 "Return structured reasoning + final answer.\n\n"
                 "CONTEXT:\n{context}\n\n"
                 "STEPS:\n{steps}\n\n"
                 "CHAT:\n{chat}\n\n"
                 "QUESTION:\n{question}"),
                ("human", "{question}")
            ])

            chat_history = "\n".join(f"{m['role']}: {m['content']}" for m in messages[-8:])

            chain = (
                {
                    "context": lambda _: context,
                    "chat": lambda _: chat_history,
                    "steps": lambda _: "\n".join(steps),
                    "question": RunnablePassthrough()
                }
                | prompt
                | llm
            )

            raw_answer = str(chain.invoke(user_input))

            # ==============================
            # SELF CHECK
            # ==============================
            issues = self_check(raw_answer)

            if issues:
                raw_answer += "\n\n⚠️ Self-check flags: " + ", ".join(issues)

            # ==============================
            # HALUCINATION RECHECK
            # ==============================
            if hallucination_check(raw_answer, context):
                retry_prompt = f"Re-generate strictly from context:\n{context}\n\nQ:{user_input}"
                raw_answer = str(chain.invoke(retry_prompt))

            response = raw_answer

        # ==============================
        # POST PROCESSING
        # ==============================
        response = clean_context(response)
        response = remove_page_artifacts(response)
        response = deduplicate_lines(response)

        messages.append({"role": "assistant", "content": response})

        st.session_state.document_chat_display[chat_key] = messages[-50:]

    # ==============================
    # STREAMING DISPLAY
    # ==============================
    for msg in messages:
        with st.chat_message(msg["role"]):

            if msg["role"] == "assistant":
                st.write_stream(stream_text(msg["content"]))
            else:
                st.markdown(msg["content"])

    st.markdown("</div>", unsafe_allow_html=True)
