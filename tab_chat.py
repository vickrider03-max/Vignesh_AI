# Auto-generated from legacy_app.py during modular refactor.
# The original monolith is retained as rollback documentation.

from functions import *
from tab_memory import get_tab_uploaded_files

from functools import lru_cache
import re
import html
import time

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
# RAG V2 SUPPORT LAYER
# ==============================

@lru_cache(maxsize=32)
def cached_vectorstore(files_key: str):
    file_names = [name for name in str(files_key or "").split("|") if name]
    if not file_names:
        return None
    return get_workspace_vector_store(file_names) or get_combined_vector_store(file_names)


def rerank_docs(query, docs):
    query_terms = {
        word
        for word in re.findall(r"\w+", str(query or "").lower())
        if len(word) > 2
    }

    def score_doc(doc):
        text = str(getattr(doc, "page_content", doc) or "").lower()
        return sum(1 for term in query_terms if term in text)

    return sorted(docs or [], key=score_doc, reverse=True)[:6]


def hallucination_check(answer: str, context: str) -> bool:
    answer_text = str(answer or "").strip()
    context_text = str(context or "").lower()
    if len(answer_text) < 40:
        return True
    if any(marker in answer_text.lower() for marker in ["i don't know", "not provided", "insufficient context"]):
        return False
    answer_terms = [
        word
        for word in re.findall(r"\w+", answer_text.lower())
        if len(word) > 4
    ][:20]
    if not answer_terms or not context_text:
        return True
    overlap_count = sum(1 for word in answer_terms if word in context_text)
    return overlap_count < max(2, len(answer_terms) // 5)


def extract_pin_diagram(text: str):
    tables = re.findall(r"(?is)(pin|connector|signal).*?\n(.*?)(?:\n\n|$)", str(text or ""))
    return tables[:3]


def reasoning_steps(query, docs):
    return [
        f"1. Understanding query: {query}",
        f"2. Retrieved {len(docs or [])} relevant chunks",
        "3. Filtering noise, headers, TOC, and OCR artifacts",
        "4. Aligning evidence across selected documents and memory",
        "5. Generating grounded answer",
        "6. Running self-check and repair if needed",
    ]


def self_check(answer):
    issues = []
    answer_text = str(answer or "")
    lower_answer = answer_text.lower()
    if "table of contents" in lower_answer:
        issues.append("Contains TOC noise")
    if len(answer_text.split()) < 30:
        issues.append("Answer too short")
    if "page_content" in lower_answer or "metadata" in lower_answer:
        issues.append("Exposes internal retrieval metadata")
    return issues


def stream_text(text: str):
    for word in str(text or "").split():
        yield word + " "
        time.sleep(0.01)


def count_text_occurrences(text, query):
    if not query:
        return 0
    return len(re.findall(re.escape(query), str(text or ""), flags=re.IGNORECASE))


def find_text_snippets(text, query, limit=5):
    snippets = []
    if not query:
        return snippets
    query_lower = query.lower()
    for line_number, line in enumerate(str(text or "").splitlines(), start=1):
        clean_line = normalize_synthesis_text(line)
        if query_lower in clean_line.lower():
            snippets.append(f"- Line {line_number}: {clean_line[:300]}")
        if len(snippets) >= limit:
            break
    return snippets


def safe_text(value, max_chars=None):
    try:
        if hasattr(value, "content"):
            text = str(value.content)
        else:
            text = str(value or "")
    except Exception:
        text = ""
    if max_chars is not None:
        return text[:max_chars]
    return text


# ==============================
# AUTONOMOUS CHAT AGENT LAYER
# ==============================

class AgentState:
    def __init__(self):
        self.plan = []
        self.context = ""
        self.memory = ""
        self.draft = ""
        self.final = ""
        self.issues = []
        self.docs = []


def _llm_text(result):
    return safe_text(result)


def _invoke_llm(llm, prompt):
    if llm is None:
        return ""
    try:
        invoke = getattr(llm, "invoke", None)
    except Exception:
        return ""
    if not callable(invoke):
        return ""
    try:
        return _llm_text(invoke(safe_text(prompt))).strip()
    except Exception:
        return ""


def _fallback_plan(query):
    text = safe_text(query).strip()
    steps = [
        "1. Identify the user's exact question and required evidence.",
        "2. Retrieve relevant document passages and workspace memory.",
        "3. Synthesize only grounded facts into a concise answer.",
        "4. Verify that the answer is supported by the retrieved context.",
    ]
    if any(word in text.lower() for word in ["compare", "difference", "versus", "vs"]):
        steps.insert(2, "3. Separate matching and differing facts across selected documents.")
    return steps


def planner_agent(llm, query):
    query_text = safe_text(query, max_chars=4000)
    prompt = "\n".join([
        "You are a planning agent.",
        "",
        "Break the query into steps.",
        "",
        "QUERY:",
        query_text,
        "",
        "Return structured steps:",
        "1.",
        "2.",
        "3.",
    ])
    plan_text = _invoke_llm(llm, prompt)
    if not plan_text:
        return _fallback_plan(query)
    return [line.strip() for line in plan_text.splitlines() if line.strip()]


def _retrieve_docs_from_vector_store(vector_store, query, limit=6):
    if vector_store is None:
        return []
    try:
        return vector_store.similarity_search(query, k=limit)
    except Exception:
        pass
    try:
        retriever = vector_store.as_retriever(search_kwargs={"k": limit})
        if hasattr(retriever, "invoke"):
            return retriever.invoke(query)
        return retriever.get_relevant_documents(query)
    except Exception:
        return []


def build_agent_memory_context(query, chat_files=None, max_chars=12000):
    memory_parts = []

    try:
        ensure_workspace_memory_loaded()
    except Exception:
        if "workspace_memory" not in st.session_state:
            st.session_state.workspace_memory = default_workspace_memory()

    try:
        memory_hits = search_workspace_memory(query, limit=4)
        for hit in memory_hits:
            memory_parts.append(f"[FAISS MEMORY]\n{hit}")
    except Exception:
        pass

    try:
        memory_text = build_unified_memory_text(
            file_names=chat_files,
            include_chat=True,
            include_agents=True,
            max_chars=max_chars,
        )
        if memory_text:
            memory_parts.append(f"[WORKSPACE MEMORY]\n{memory_text}")
    except Exception:
        pass

    try:
        user_id = get_active_user_id()
        for entry in get_chatpdf_memory(user_id, chat_files or [])[-8:]:
            memory_parts.append(
                "[DOCUMENT CHAT MEMORY]\n"
                f"Question: {entry.get('question', '')}\n"
                f"Answer: {entry.get('answer', '')}"
            )
    except Exception:
        pass

    return clean_context("\n\n".join(memory_parts))[:max_chars]


def retriever_agent(query, chat_files):
    files_key = "|".join(chat_files or [])
    doc_docs = _retrieve_docs_from_vector_store(cached_vectorstore(files_key), query, limit=8)
    if not doc_docs:
        doc_docs = _retrieve_docs_from_vector_store(get_combined_vector_store(chat_files), query, limit=8)
    doc_docs = rerank_docs(query, doc_docs)

    memory_docs = _retrieve_docs_from_vector_store(get_workspace_vector_store(chat_files), query, limit=4)
    docs = rerank_docs(query, doc_docs + memory_docs)

    if not docs:
        try:
            docs = sparse_chatpdf_search(query, chat_files, user_id=get_active_user_id(), top_k=8)
        except Exception:
            docs = []
    docs = rerank_docs(query, docs)

    doc_context = "\n\n".join(
        f"[DOCUMENT CONTEXT]\n{getattr(doc, 'page_content', str(doc))}"
        for doc in docs
    )
    memory_context = build_agent_memory_context(query, chat_files)
    context = clean_context(doc_context + "\n\n" + memory_context)
    return context, memory_context, docs


def _fallback_reasoning_answer(query, context):
    sentences = document_intelligence_meaningful_sentences(context, limit=6)
    if not sentences:
        return (
            "I could not find enough grounded context in the selected documents or memory to answer "
            "that accurately. Try selecting the relevant file or asking for a more specific section."
        )
    answer_lines = [normalize_synthesis_text(sentence)[:420] for sentence in sentences if sentence]
    unique_lines = []
    for line in answer_lines:
        if line and line not in unique_lines:
            unique_lines.append(line)
    return "Based on the retrieved document and memory context:\n\n" + "\n".join(
        f"- {line}" for line in unique_lines[:5]
    )


def reasoning_agent(llm, query, context, chat_history):
    prompt = "\n".join([
        "You are a reasoning agent.",
        "",
        "Use ONLY the context. Think step-by-step internally, but do not reveal hidden reasoning.",
        "Produce a structured, grounded answer. If context is insufficient, say so clearly.",
        "",
        "CONTEXT:",
        safe_text(context, max_chars=60000),
        "",
        "CHAT HISTORY:",
        safe_text(chat_history, max_chars=12000),
        "",
        "QUESTION:",
        safe_text(query, max_chars=4000),
    ])
    answer = _invoke_llm(llm, prompt)
    return answer or _fallback_reasoning_answer(query, context)


def _answer_looks_weak(answer):
    text = str(answer or "").strip().lower()
    weak_markers = [
        "i don't know",
        "i do not know",
        "not enough",
        "could not find",
        "insufficient",
        "no relevant",
    ]
    return len(text) < 80 or any(marker in text for marker in weak_markers)


def verification_agent(llm, query, answer, context):
    prompt = "\n".join([
        "You are a verification agent.",
        "",
        "Check if the answer is grounded in the context.",
        "",
        "QUESTION:",
        safe_text(query, max_chars=4000),
        "",
        "ANSWER:",
        safe_text(answer, max_chars=16000),
        "",
        "CONTEXT:",
        safe_text(context, max_chars=60000),
        "",
        "Return:",
        "- OK or ISSUE",
        "- list problems if any",
    ])
    result = _invoke_llm(llm, prompt)
    issues = []
    if result and "ISSUE" in result.upper():
        issues.append(result)
    if _answer_looks_weak(answer):
        issues.append("ISSUE: The answer appears weak, incomplete, or under-supported.")
    if not str(context or "").strip():
        issues.append("ISSUE: No retrieval context was available.")
    issues.extend(f"ISSUE: {issue}" for issue in self_check(answer))
    if hallucination_check(answer, context):
        issues.append("ISSUE: The answer may not be sufficiently grounded in retrieved context.")
    return issues


def _repair_answer(llm, query, answer, context, issues):
    repair_prompt = "\n".join([
        "Fix the answer using context only.",
        "",
        "ISSUES:",
        safe_text(issues, max_chars=8000),
        "",
        "ANSWER:",
        safe_text(answer, max_chars=16000),
        "",
        "CONTEXT:",
        safe_text(context, max_chars=60000),
        "",
        "QUESTION:",
        safe_text(query, max_chars=4000),
    ])
    repaired = _invoke_llm(llm, repair_prompt)
    if repaired:
        return repaired
    return _fallback_reasoning_answer(query, context)


def autonomous_agent_run(llm, query, chat_files, chat_history):
    state = AgentState()

    state.plan = planner_agent(llm, query)
    state.context, state.memory, state.docs = retriever_agent(query, chat_files)
    state.plan.extend(reasoning_steps(query, state.docs))
    state.draft = reasoning_agent(llm, query, state.context, chat_history)
    state.issues = verification_agent(llm, query, state.draft, state.context)

    if state.issues:
        state.final = _repair_answer(llm, query, state.draft, state.context, state.issues)
        state.issues = verification_agent(llm, query, state.final, state.context)
    else:
        state.final = state.draft

    if state.docs and "Sources:" not in str(state.final):
        state.final = str(state.final).rstrip() + "\n\nSources:\n" + format_chatpdf_sources(state.docs)

    pin_diagrams = extract_pin_diagram(state.context)
    if pin_diagrams and any(term in str(query or "").lower() for term in ["pin", "connector", "signal", "diagram"]):
        pin_lines = []
        for label, table in pin_diagrams:
            pin_lines.append(f"- {normalize_synthesis_text(label + ': ' + table)[:500]}")
        state.final = str(state.final).rstrip() + "\n\nRelevant pin/connector evidence:\n" + "\n".join(pin_lines)

    st.session_state.chat_agent_trace = {
        "plan": state.plan,
        "issues": state.issues,
        "context_chars": len(state.context or ""),
        "memory_chars": len(state.memory or ""),
        "doc_count": len(state.docs or []),
    }
    return state.final


def save_memory(user_input, response, chat_files=None):
    try:
        append_chatpdf_memory(get_active_user_id(), chat_files, user_input, response)
    except Exception:
        pass

    try:
        ensure_workspace_memory_loaded()
        append_chat_to_workspace_memory(user_input, response, chat_files)
        save_workspace_memory()
        save_memory_log(
            "chat_agent",
            "Stored autonomous chat exchange in workspace memory.",
            {"files": list(chat_files or [])},
        )
    except Exception:
        pass


def show_agent_trace(state):
    st.sidebar.subheader("Agent Trace")
    st.sidebar.write("### Plan")
    st.sidebar.write(state.plan if isinstance(state, AgentState) else state.get("plan", []))
    st.sidebar.write("### Issues")
    st.sidebar.write(state.issues if isinstance(state, AgentState) else state.get("issues", []))


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

    current_chat_messages = []
    chat_display_key = None
    available_chat_files = list(dict.fromkeys(st.session_state.get("selected_files", [])))

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
                        count = count_text_occurrences(combined_text, word)
                        response = f"Found '{word}' {count} time(s) in the selected document text."
                    else:
                        response = "Specify quoted text."

                elif is_find_query:
                    match = re.search(r"'(.*?)'|\"(.*?)\"", processing_input)
                    if match:
                        query = match.group(1) or match.group(2)
                        snippets = find_text_snippets(combined_text, query)
                        if snippets:
                            response = f"Search results for '{query}':\n\n" + "\n".join(snippets)
                        else:
                            response = f"No direct text matches found for '{query}' in the selected documents."
                    else:
                        response = "Specify quoted search text."

                # ==============================
                # RAG PIPELINE (FIXED)
                # ==============================
                else:
                    llm = load_llm()

                    chat_history = "\n".join(
                        f"{m['role']}: {m['content']}"
                        for m in current_chat_messages[-10:]
                    )

                    response = autonomous_agent_run(
                        llm,
                        processing_input,
                        chat_files,
                        chat_history
                    )

                # ==============================
                # POST PROCESSING (FIXED)
                # ==============================
                response = clean_context(response)
                response = remove_page_artifacts(response)
                response = deduplicate_lines(response)

                current_chat_messages.append({"role": "assistant", "content": response})

                st.session_state.document_chat_display[chat_display_key] = current_chat_messages[-50:]

                st.session_state.messages = current_chat_messages
                save_memory(user_input, response, chat_files)
                st.session_state.chat_stream_target = (chat_display_key, len(current_chat_messages) - 1)

    # ==============================
    # DISPLAY CHAT
    # ==============================
    visible_messages = st.container()
    with visible_messages:
        for message_index, message in enumerate(current_chat_messages):
            with st.chat_message(message["role"]):
                should_stream = (
                    message["role"] == "assistant"
                    and hasattr(st, "write_stream")
                    and chat_display_key is not None
                    and st.session_state.get("chat_stream_target") == (chat_display_key, message_index)
                )
                if should_stream:
                    st.write_stream(stream_text(message["content"]))
                    st.session_state.chat_stream_target = None
                else:
                    st.markdown(message["content"])

    st.markdown("</div>", unsafe_allow_html=True)
