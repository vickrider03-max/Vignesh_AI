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
# AGENT STATE (FIXED - IMPORTANT)
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
# CACHE VECTOR STORE
# ==============================

@lru_cache(maxsize=32)
def cached_vectorstore(files_key: str):
    file_names = [name for name in str(files_key or "").split("|") if name]
    if not file_names:
        return None
    return get_workspace_vector_store(file_names) or get_combined_vector_store(file_names)


# ==============================
# RETRIEVAL CORE
# ==============================

def rerank_docs(query, docs):
    query_terms = {w for w in re.findall(r"\w+", str(query).lower()) if len(w) > 2}

    def score(doc):
        text = str(getattr(doc, "page_content", doc)).lower()
        return sum(1 for t in query_terms if t in text)

    return sorted(docs or [], key=score, reverse=True)[:6]


def _retrieve(vector_store, query, k=6):
    if not vector_store:
        return []
    try:
        return vector_store.similarity_search(query, k=k)
    except:
        try:
            return vector_store.as_retriever(search_kwargs={"k": k}).invoke(query)
        except:
            return []


# ==============================
# SIMPLE LLM WRAPPER
# ==============================

def _invoke_llm(llm, prompt):
    if not llm:
        return ""
    try:
        return str(llm.invoke(prompt))
    except:
        return ""


# ==============================
# RAG PIPELINE
# ==============================

def retriever_agent(query, chat_files):
    key = "|".join(chat_files or [])

    docs = _retrieve(cached_vectorstore(key), query, k=8)
    if not docs:
        docs = _retrieve(get_combined_vector_store(chat_files), query, k=8)

    docs = rerank_docs(query, docs)

    context = "\n\n".join(
        str(getattr(d, "page_content", d)) for d in docs
    )

    return clean_context(context), docs


# ==============================
# AGENT CORE
# ==============================

def reasoning_agent(llm, query, context, chat_history):
    prompt = f"""
You are a strict RAG assistant.

RULES:
- Use ONLY context
- If unsure, say "not found in document"

CONTEXT:
{context}

CHAT:
{chat_history}

QUESTION:
{query}
"""
    return _invoke_llm(llm, prompt)


def verification(answer, context):
    if len(answer or "") < 30:
        return ["Answer too short"]
    if "not found" in answer.lower():
        return []
    if len(set(answer.lower().split()) & set(context.lower().split())) < 5:
        return ["Low grounding in context"]
    return []


def repair(llm, query, answer, context, issues):
    prompt = f"""
Fix this answer using ONLY context.

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


def autonomous_agent_run(llm, query, chat_files, chat_history):
    state = AgentState()

    state.context, state.docs = retriever_agent(query, chat_files)

    state.draft = reasoning_agent(llm, query, state.context, chat_history)

    state.issues = verification(state.draft, state.context)

    if state.issues:
        state.final = repair(llm, query, state.draft, state.context, state.issues)
    else:
        state.final = state.draft

    if not state.final:
        state.final = "No valid answer generated from document context."

    return state.final


# ==============================
# CHAT TAB UI
# ==============================

def render_chat_tab():
    st.markdown('<div id="chat-section">', unsafe_allow_html=True)

    if "document_chat_display" not in st.session_state:
        st.session_state.document_chat_display = {}

    available_chat_files = list(dict.fromkeys(st.session_state.get("selected_files", [])))

    if not available_chat_files:
        st.warning("Upload files first")
        return

    chat_files = st.multiselect(
        "Choose file(s) for Chat",
        options=available_chat_files,
        default=st.session_state.get("chat_file_selection", [])
    )

    if not chat_files:
        return

    chat_key = "|".join(chat_files)
    messages = st.session_state.document_chat_display.setdefault(chat_key, [])

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

        st.session_state.document_chat_display[chat_key] = messages[-50:]

    # display
    for msg in messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
