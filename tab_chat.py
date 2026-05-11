# ==============================
# FIXED + RESTORED VERSION
# ==============================

from functions import *
from tab_memory import get_tab_uploaded_files
from functools import lru_cache
import re
import html
import time

# ==============================
# HELPERS (UNCHANGED + SAFE)
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
# 🔥 ONLY ADDITION (SAFE FIX)
# ==============================

def is_valid_context(context: str) -> bool:
    if not context:
        return False

    text = context.lower()

    bad_signals = [
        "uploaded and queued",
        "processing",
        "loading",
        "no content",
        "empty",
    ]

    # must contain real semantic content
    words = [w for w in text.split() if len(w) > 3]
    has_signal = len(words) > 40

    return has_signal and not any(b in text for b in bad_signals)


# ==============================
# RAG + AGENT SYSTEM (UNCHANGED)
# ==============================

@lru_cache(maxsize=32)
def cached_vectorstore(files_key: str):
    file_names = [name for name in str(files_key or "").split("|") if name]
    if not file_names:
        return None
    return get_workspace_vector_store(file_names) or get_combined_vector_store(file_names)


def rerank_docs(query, docs):
    query_terms = {
        word for word in re.findall(r"\w+", str(query or "").lower())
        if len(word) > 2
    }

    def score(doc):
        text = str(getattr(doc, "page_content", doc)).lower()
        return sum(1 for t in query_terms if t in text)

    return sorted(docs or [], key=score, reverse=True)[:6]


def hallucination_check(answer: str, context: str) -> bool:
    if not answer or not context:
        return True

    answer = answer.lower()
    context = context.lower()

    if len(answer) < 50:
        return True

    terms = [w for w in re.findall(r"\w+", answer) if len(w) > 4][:20]
    overlap = sum(1 for t in terms if t in context)

    return overlap < 2


# ==============================
# RETRIEVER (FIX ONLY HERE)
# ==============================

def retriever_agent(query, chat_files):

    files_key = "|".join(chat_files or [])

    doc_docs = _retrieve_docs_from_vector_store(
        cached_vectorstore(files_key), query, limit=8
    )

    if not doc_docs:
        doc_docs = _retrieve_docs_from_vector_store(
            get_combined_vector_store(chat_files), query, limit=8
        )

    doc_docs = rerank_docs(query, doc_docs)

    memory_docs = _retrieve_docs_from_vector_store(
        get_workspace_vector_store(chat_files), query, limit=4
    )

    docs = rerank_docs(query, doc_docs + memory_docs)

    doc_context = "\n\n".join(
        f"[DOC]\n{getattr(d, 'page_content', str(d))}"
        for d in docs
    )

    memory_context = build_agent_memory_context(query, chat_files)
    context = clean_context(doc_context + "\n\n" + memory_context)

    # 🔥 ONLY FIX HERE
    if not is_valid_context(context):
        context = ""

    return context, memory_context, docs


# ==============================
# REASONING (RESTORED)
# ==============================

def reasoning_agent(llm, query, context, chat_history):

    if not is_valid_context(context):
        return "No sufficient document context found to answer this query."

    prompt = f"""
You are a reasoning engine.

RULES:
- Use ONLY context
- Do not hallucinate
- If unsure, say so

CONTEXT:
{safe_text(context, 40000)}

CHAT:
{safe_text(chat_history, 8000)}

QUESTION:
{safe_text(query, 3000)}
"""

    answer = _invoke_llm(llm, prompt)

    if hallucination_check(answer, context):
        return "The document does not contain enough grounded information."

    return answer


# ==============================
# AUTONOMOUS AGENT (FULL RESTORED)
# ==============================

def autonomous_agent_run(llm, query, chat_files, chat_history):

    state = AgentState()

    state.plan = planner_agent(llm, query)
    state.context, state.memory, state.docs = retriever_agent(query, chat_files)

    state.plan.extend(reasoning_steps(query, state.docs))

    if not state.context:
        return "No relevant document content found."

    state.draft = reasoning_agent(llm, query, state.context, chat_history)

    state.issues = verification_agent(llm, query, state.draft, state.context)

    if state.issues:
        state.final = _repair_answer(
            llm, query, state.draft, state.context, state.issues
        )
    else:
        state.final = state.draft

    if state.docs and "Sources:" not in state.final:
        state.final += "\n\nSources:\n" + format_chatpdf_sources(state.docs)

    pin_diagrams = extract_pin_diagram(state.context)
    if pin_diagrams:
        state.final += "\n\nPin/Signal Info:\n" + str(pin_diagrams)

    st.session_state.chat_agent_trace = {
        "plan": state.plan,
        "issues": state.issues,
        "doc_count": len(state.docs),
        "context_chars": len(state.context or "")
    }

    return state.final


# ==============================
# CHAT TAB (UNCHANGED STRUCTURE)
# ==============================

def render_chat_tab():

    current_chat_messages = []

    available_chat_files = list(dict.fromkeys(
        st.session_state.get("selected_files", [])
    ))

    if available_chat_files:

        chat_files = st.multiselect(
            "Choose file(s) for Chat",
            options=available_chat_files,
            default=st.session_state.get("chat_file_selection", [])
        )

        if chat_files:

            user_input = st.chat_input("Ask anything from documents")

            if user_input:

                current_chat_messages.append({"role": "user", "content": user_input})

                llm = load_llm()

                chat_history = "\n".join(
                    f"{m['role']}: {m['content']}"
                    for m in current_chat_messages[-10:]
                )

                response = autonomous_agent_run(
                    llm,
                    user_input,
                    chat_files,
                    chat_history
                )

                # 🔥 FINAL SAFETY FIX ONLY
                if not response or "uploaded and queued" in response.lower():
                    response = "No valid grounded answer could be generated."

                response = clean_context(response)
                response = remove_page_artifacts(response)
                response = deduplicate_lines(response)

                current_chat_messages.append(
                    {"role": "assistant", "content": response}
                )

                st.session_state.messages = current_chat_messages

    for msg in current_chat_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
