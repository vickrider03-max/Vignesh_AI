from functions import *
from tab_memory import get_tab_uploaded_files
from functools import lru_cache
import re
import time

# ==============================
# CLEANING HELPERS
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
        k = line.strip().lower()
        if k and k not in seen:
            seen.add(k)
            out.append(line)
    return "\n".join(out)


def remove_noise(text: str) -> str:
    # IMPORTANT FIX: remove memory spam responses
    if not text:
        return ""

    bad_markers = [
        "[MEMORY EVENT]",
        "queued for processing",
        "Uploaded and queued"
    ]

    for m in bad_markers:
        text = text.replace(m, "")

    return text.strip()


# ==============================
# VECTOR CACHE (FIXED)
# ==============================

@lru_cache(maxsize=32)
def cached_vectorstore(files_key: str):
    files = [f for f in files_key.split("|") if f]
    if not files:
        return None
    return get_workspace_vector_store(files) or get_combined_vector_store(files)


def safe_retrieve(vector_store, query, k=6):
    if not vector_store:
        return []
    try:
        return vector_store.similarity_search(query, k=k)
    except:
        try:
            retriever = vector_store.as_retriever(search_kwargs={"k": k})
            return retriever.invoke(query)
        except:
            return []


def rerank(query, docs):
    if not docs:
        return []

    q = set(re.findall(r"\w+", query.lower()))
    def score(d):
        t = str(getattr(d, "page_content", d)).lower()
        return sum(1 for w in q if w in t)

    return sorted(docs, key=score, reverse=True)[:5]


# ==============================
# CORE RAG PIPELINE (FIXED)
# ==============================

def build_context(query, files):
    key = "|".join(files)

    docs = safe_retrieve(cached_vectorstore(key), query, 6)
    if not docs:
        docs = safe_retrieve(get_combined_vector_store(files), query, 6)

    docs = rerank(query, docs)

    doc_text = "\n\n".join(
        str(getattr(d, "page_content", d)) for d in docs
    )

    doc_text = clean_context(doc_text)

    return doc_text, docs


def generate_answer(llm, query, context):
    if not llm:
        return "LLM not loaded."

    prompt = f"""
You are a document QA assistant.

RULES:
- Answer ONLY using context
- If not found, say "Not found in document"
- Do NOT include system logs or memory events

CONTEXT:
{context}

QUESTION:
{query}
"""

    try:
        return llm.invoke(prompt)
    except:
        return "Error generating response."


# ==============================
# STREAMLIT UI
# ==============================

def render_chat_tab():
    st.markdown('<div id="chat-section">', unsafe_allow_html=True)

    if "document_chat_display" not in st.session_state:
        st.session_state.document_chat_display = {}

    available_files = list(dict.fromkeys(st.session_state.get("selected_files", [])))

    if not available_files:
        st.warning("No files selected.")
        return

    chat_files = st.multiselect(
        "Choose file(s)",
        options=available_files,
        default=st.session_state.get("chat_file_selection", [])
    )

    if not chat_files:
        return

    chat_key = get_chatpdf_memory_key(get_active_user_id(), chat_files)
    messages = st.session_state.document_chat_display.setdefault(chat_key, [])

    user_input = st.chat_input("Ask something from document")

    if user_input:

        messages.append({"role": "user", "content": user_input})

        llm = load_llm()

        # ==========================
        # STEP 1: RETRIEVE CONTEXT
        # ==========================
        context, docs = build_context(user_input, chat_files)

        # IMPORTANT FIX: prevent memory-only answers
        if not context.strip():
            response = "No relevant information found in the document."
        else:
            # ==========================
            # STEP 2: GENERATE ANSWER
            # ==========================
            response = generate_answer(llm, user_input, context)

        # cleanup hallucinated memory logs
        response = remove_noise(response)
        response = deduplicate_lines(response)

        messages.append({"role": "assistant", "content": response})

        st.session_state.document_chat_display[chat_key] = messages[-50:]

    # ==============================
    # DISPLAY
    # ==============================
    for m in messages:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    st.markdown("</div>", unsafe_allow_html=True)
