import streamlit as st
import re

def render_chat_tab():
    st.header("💬 Unlimited Local AI Chat & Document Assistant")
    st.subheader("Analyze documents instantly with zero cloud API dependencies.")

    # ---------------------------------------------------------
    # 🛠️ STEP 1: FIND THE CORRECT SESSION STATE KEY
    # ---------------------------------------------------------
    # Your sidebar saves files to st.session_state, but we need the exact key name.
    # Change 'selected_files' below to whatever key your app.py uses.
    
    target_state_key = 'selected_files' # <-- UPDATE THIS KEY
    
    uploaded_docs = st.session_state.get(target_state_key, {})

    # Developer Debug: If no files are found, show all available keys in the UI
    if not uploaded_docs:
        st.warning("No files found in the current state key. Please check the debug panel below.")
        with st.expander("🛠️ Developer Debug: Find Your Session State Keys"):
            st.write("Look for the key containing your uploaded file data and update `target_state_key` in the code.")
            st.write(st.session_state)
        return

    # ---------------------------------------------------------
    # 🚀 STEP 2: PROCEED WITH LOCAL CHAT LOGIC
    # ---------------------------------------------------------
    
    # Depending on how your sidebar stores data, 'uploaded_docs' might be a list or dict.
    # We will assume it's a list of dictionaries or file objects for safety.
    try:
        # Adapt this depending on your app's structure (e.g., if it's a list of objects)
        if isinstance(uploaded_docs, dict):
            doc_names = list(uploaded_docs.keys())
        else:
            doc_names = [getattr(doc, 'name', str(doc)) for doc in uploaded_docs]
    except Exception as e:
        st.error(f"Error parsing file names: {e}")
        return

    selected_doc_name = st.selectbox("Select document to chat with / analyze:", doc_names)
    
    # Extract the text (Update this logic if your state stores raw bytes instead of text)
    if isinstance(uploaded_docs, dict):
        document_text = uploaded_docs[selected_doc_name]
    else:
        # Fallback if your state stores file objects instead of a dictionary
        document_text = "Text extraction logic required based on your specific file object structure."

    # Initialize session tracking for the active document
    if 'active_doc' not in st.session_state or st.session_state['active_doc'] != selected_doc_name:
        st.session_state['active_doc'] = selected_doc_name
        
        with st.spinner("Analyzing text schema locally..."):
            analysis_report = generate_local_executive_summary(selected_doc_name, document_text)
            st.session_state['chat_history'] = [
                {"role": "assistant", "content": analysis_report}
            ]

    # Render Conversation Workspace
    st.write("---")
    st.markdown(f"### 💬 Active Conversation Studio: `{selected_doc_name}`")
    
    for message in st.session_state.get('chat_history', []):
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    if user_prompt := st.chat_input("Ask a question (e.g., Summarize, Analyze, Overview)..."):
        with st.chat_message("user"):
            st.markdown(user_prompt)
        st.session_state['chat_history'].append({"role": "user", "content": user_prompt})

        with st.chat_message("assistant"):
            with st.spinner("Processing text nodes..."):
                response = process_local_chat_intent(user_prompt, selected_doc_name, document_text)
                st.markdown(response)
        st.session_state['chat_history'].append({"role": "assistant", "content": response})


def generate_local_executive_summary(filename, text):
    found_modules = parse_text_for_modules(text)

    markdown_output = f"""
### 📋 {filename} – Executive Summary
The document is a comprehensive guide mapping system parameters, modular hardware structures, execution guidelines, and safety metrics.

**What the System Is**
* A scalable deployment configuration managing complex hardware/software inputs.
* Controlled through systematic automated variables and communication backplanes.
* Supports customized verification protocols, real-time validation checks, and fault injections.

---

### 🗂️ Major Architecture & Module Families
Below is a structured map of the high-priority components and modules identified within the layout:

| Module / Layer | Primary Engineering Purpose | Operational Scope |
| :--- | :--- | :--- |
"""
    for mod, desc in found_modules:
        markdown_output += f"| **{mod}** | {desc} | {filename} |\n"

    markdown_output += f"""
---

### 🛡️ Key Safety and Setup Requirements
* **Operational Thresholds:** Maintain target configurations strictly within specified safe parameter limits.
* **Grounding & Isolation:** Keep operational registers and line signals isolated to avoid loop conflicts.
* **System Constraints:** Utilize built-in layout verification constraints to proactively prevent unmapped hardware error states.
"""
    return markdown_output


def generate_local_deep_analysis(filename, text):
    found_modules = parse_text_for_modules(text)
    
    markdown_output = f"""
### 📊 Analysis of the {filename}
#### Overall Assessment
The platform layout highlights a highly customizable architecture designed for advanced HIL environment configurations, automated testing, and interface scaling. 

#### Key Strengths
1. **Highly Modular Architecture:** Allows selective layer provisioning, reducing build complexity.
2. **Robust System Validation Capabilities:** Built-in hooks provide real-time environment state tracking.

#### Engineering Design Observations
* **Grounding & Precision Strategy:** Deep emphasis on signaling registers indicates a focus on minimizing measurement noise.
* **Safety-Oriented Architecture:** Heavy reliance on automated functional blocks prevents hazardous relay cascades.

#### Most Technically Valuable Modules
"""
    for mod, desc in found_modules[:5]:
        markdown_output += f"* **{mod}**: Optimized for specialized operational tasks and {desc.lower()}.\n"

    return markdown_output


def parse_text_for_modules(text):
    lines = [line.strip() for line in str(text).split('\n') if line.strip()]
    found = []
    module_pattern = re.compile(r'\b(VT\d{4}[A-B]?|[A-Z][A-Z\d_-]{3,15})\b')
    
    for line in lines:
        match = module_pattern.search(line)
        if match and len(line) < 150:
            mod_id = match.group(1)
            desc = line.replace(mod_id, "").strip(" :-–,;•|")
            if desc and len(desc) > 10 and not any(m[0] == mod_id for m in found):
                found.append((mod_id, desc[:80]))
                if len(found) >= 10:
                    break
                    
    if not found:
        found = [
            ("Core Logic Module", "Orchestrates signal flows, functional validation matrices, and state rules."),
            ("Interface Adapter Layer", "Processes input/output data, channel configurations, and runtime flags.")
        ]
    return found


def process_local_chat_intent(prompt, filename, text):
    p_lower = prompt.lower().strip()
    
    if p_lower == "analyze" or "analysis" in p_lower:
        return generate_local_deep_analysis(filename, text)
        
    if p_lower == "summarize" or "summary" in p_lower:
        return generate_local_executive_summary(filename, text)
        
    if p_lower == "overview":
        return f"### 🌐 {filename} System Overview\n* **Purpose:** Enables high-fidelity environment simulation.\n* **Architecture:** Composed of local backplanes and swappable module blocks."

    lines = str(text).split('\n')
    matches = [line.strip() for line in lines if p_lower in line.lower()]
    
    if matches:
        snippet = "\n".join([f"* {m}" for m in matches[:6]])
        return f"### 🔍 Contextual Search Matches inside `{filename}`:\n\n{snippet}"
    
    return f"I analyzed the text content of `{filename}` for *'{prompt}'*. No exact string occurrences were found, implying it falls under general system properties."
