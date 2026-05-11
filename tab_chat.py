from functions import *
from tab_memory import get_tab_uploaded_files

from functools import lru_cache
import re
import html
import time
import streamlit as st

from langchain_core.prompts import ChatPromptTemplate
from langchain_core.runnables import RunnablePassthrough


# ==============================
# HELPERS
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
# AGENT STATE
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


# ==============================
# VECTOR CACHE
# ==============================

@lru_cache(maxsize=32)
def cached_vectorstore(files_key: str):
    file_names = [f for f in str(files_key or "").split("|") if f]
    if not file_names:
        return None
    return get_workspace_vector_store(file_names) or get_combined_vector_store(file_names)


# ==============================
# RERANK (simple but stable)
# ==============================

def rerank_docs(query, docs):
    query_terms = {w for w in re.findall(r"\w+", str(query).lower()) if len(w) > 2}

    def score(doc):
        text = str(getattr(doc, "page_content", doc)).lower()
        return sum(1 for t in query_terms if t in text)

    return sorted(docs or [], key=score, reverse=True)[:6]


# ==============================
# SAFE LLM CALL
# ==============================

def _invoke_llm(llm, prompt):
    if not llm:
        return ""
    try:
        return str(llm.invoke(prompt))
    except:
        return ""


# ==============================
# 🔥 FIXED RETRIEVAL (IMPORTANT FIX)
# ==============================

def retriever_agent(query, chat_files):
    key = "|".join(chat_files or [])

    docs = []

    # 1. VECTOR DB (highest priority)
    vector_store = cached_vectorstore(key)
    if vector_store:
        try:
            docs = vector_store.similarity_search(query, k=10)
        except:
            try:
                docs = vector_store.as_retriever(search_kwargs={"k": 10}).invoke(query)
            except:
                docs = []

    # 2. COMBINED VECTOR STORE fallback
    if not docs:
        vs = get_combined_vector_store(chat_files)
        if vs:
            try:
                docs = vs.similarity_search(query, k=10)
            except:
                docs = []

    # 3. RAW TEXT fallback (CRITICAL)
    if not docs:
        raw_text = "\n".join(
            st.session_state.file_texts.get(f, "")
            for f in (chat_files or [])
        )
        docs = [raw_text[i:i+2000] for i in range(0, len(raw_text), 2000)]

    docs = rerank_docs(query, docs)

    context = "\n\n".join(
        str(getattr(d, "page_content", d)) for d in docs
    )

    return clean_context(context), docs


# ==============================
# REASONING
# ==============================

def reasoning_agent(llm, query, context, chat_history):
    prompt = f"""
You are a strict document QA system.

RULES:
- Use ONLY context
- If answer is not in context, say "Not found in document"

CONTEXT:
{context}

CHAT:
{chat_history}

QUESTION:
{query}
"""
    return _invoke_llm(llm, prompt)


# ==============================
# VERIFICATION
# ==============================

def verification(answer, context):
    issues = []
    if len(answer or "") < 30:
        issues.append("Too short")
    if len(set(answer.lower().split()) & set(context.lower().split())) < 5:
        issues.append("Weak grounding")
    return issues


def repair(llm, query, answer, context, issues):
    prompt = f"""
Fix answer using ONLY context.

ISSUES:
{issues}

ANSWER:
{answer}

CONTEXT:
{context}

QUESTION:
{query}
"""
    return _invoke_llm(llm, prompt)


# ==============================
# AGENT PIPELINE
# ==============================

def autonomous_agent_run(llm, query, chat_files, chat_history):
    state = AgentState()

    state.context, state.docs = retriever_agent(query, chat_files)

    state.draft = reasoning_agent(llm, query, state.context, chat_history)

    state.issues = verification(state.draft, state.context)

    if state.issues:
        state.final = repair(llm, query, state.draft, state.context, state.issues)
    else:
        state.final = state.draft

    # 🔥 SAFE FALLBACK (IMPORTANT FIX)
    if not state.final or len(state.final.strip()) < 20:
        if state.context.strip():
            state.final = state.context[:2500]
        else:
            state.final = "No relevant content found in documents."

    return state.final


# ==============================
# CHAT UI
# ==============================

def render_chat_tab():
    st.markdown('<div id="chat-section">', unsafe_allow_html=True)

    if "document_chat_display" not in st.session_state:
        st.session_state.document_chat_display = {}

    available_files = list(dict.fromkeys(st.session_state.get("selected_files", [])))

    if not available_files:
        st.warning("Upload files first")
        return

    chat_files = st.multiselect(
        "Choose file(s) for Chat",
        options=available_files,
        default=st.session_state.get("chat_file_selection", [])
    )

    if not chat_files:
        return

    key = "|".join(chat_files)
    messages = st.session_state.document_chat_display.setdefault(key, [])

    user_input = st.chat_input("Ask from documents")

    if user_input:
        messages.append({"role": "user", "content": user_input})

        llm = load_llm()

        chat_history = "\n".join(
            f"{m['role']}: {m['content']}" for m in messages[-10:]
        )

        response = autonomous_agent_run(
            llm,
            user_input,
            chat_files,
            chat_history
        )

        response = clean_context(response)
        response = remove_page_artifacts(response)
        response = deduplicate_lines(response)

        messages.append({"role": "assistant", "content": response})

        st.session_state.document_chat_display[key] = messages[-50:]

    for msg in messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
