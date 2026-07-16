import streamlit as st
import torch
from transformers import AutoModelForCausalLM, AutoTokenizer, pipeline

# -------------------------------------------------------------------------
# AI MODEL INITIALIZATION (Runs entirely locally)
# -------------------------------------------------------------------------
MODEL_ID = "microsoft/Phi-3-mini-4k-instruct"

@st.cache_resource
def get_ai_pipeline():
    """Initializes the local model for text generation."""
    tokenizer = AutoTokenizer.from_pretrained(MODEL_ID)
    model = AutoModelForCausalLM.from_pretrained(
        MODEL_ID, 
        device_map="cpu", 
        torch_dtype="auto", 
        trust_remote_code=True
    )
    return pipeline("text-generation", model=model, tokenizer=tokenizer)

def render_chat_tab():
    st.header("🧠 Local AI Document Intelligence")
    st.markdown("Run your private, local LLM to synthesize and rewrite document content.")

    # 1. DOCUMENT SELECTION
    target_state_key = 'selected_files' 
    uploaded_docs = st.session_state.get(target_state_key, {})

    if not uploaded_docs:
        st.info("Please upload/select documents in the sidebar first.")
        return

    doc_names = list(uploaded_docs.keys()) if isinstance(uploaded_docs, dict) else [getattr(d, 'name', str(d)) for d in uploaded_docs]
    
    selected_doc_name = st.selectbox(
        "Select document to analyze:", 
        options=doc_names, 
        index=None, 
        placeholder="Choose a document..."
    )
    
    if not selected_doc_name:
        return

    # 2. EXTRACT TEXT
    document_text = uploaded_docs[selected_doc_name] if isinstance(uploaded_docs, dict) else str(uploaded_docs[doc_names.index(selected_doc_name)])
    
    # 3. LOAD AI MODEL
    with st.spinner("Waking up Local AI..."):
        try:
            generator = get_ai_pipeline()
        except Exception as e:
            st.error(f"Could not load AI model: {e}")
            return

    # 4. CHAT HISTORY
    if 'doc_chat_histories' not in st.session_state:
        st.session_state['doc_chat_histories'] = {}

    if selected_doc_name not in st.session_state['doc_chat_histories']:
        st.session_state['doc_chat_histories'][selected_doc_name] = []

    active_history = st.session_state['doc_chat_histories'][selected_doc_name]

    # Render History
    for message in active_history:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    # 5. GENERATIVE AI INPUT
    if user_prompt := st.chat_input("Ask a question about this document..."):
        with st.chat_message("user"):
            st.markdown(user_prompt)
        
        active_history.append({"role": "user", "content": user_prompt})

        # Generate Response
        with st.chat_message("assistant"):
            with st.spinner("AI is reading and synthesizing..."):
                # We limit the context to fit model token windows
                context = f"Document: {document_text[:3500]}\n\nQuestion: {user_prompt}\n\nAnswer:"
                
                # Model generation parameters for better "rewriting" quality
                output = generator(
                    context, 
                    max_new_tokens=500, 
                    do_sample=True, 
                    temperature=0.7, 
                    top_k=50
                )
                
                # Clean up the output string
                raw_response = output[0]['generated_text']
                final_response = raw_response.split("Answer:")[-1].strip()
                
                st.markdown(final_response)
        
        active_history.append({"role": "assistant", "content": final_response})
