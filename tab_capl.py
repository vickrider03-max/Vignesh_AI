# Auto-generated from legacy_app.py during modular refactor.
# The original monolith is retained as rollback documentation.

from functions import *

# ==============================
# CAPL TAB UI
# Moved from legacy_app.py tab body.
# UI rendering and event handling live here; backend work is delegated to functions.py.
# ==============================
def render_capl_tab():
    st.markdown('<div id="capl-section">', unsafe_allow_html=True)
    capl_header_col, capl_reset_col = st.columns([8, 1])
    with capl_header_col:
        st.subheader("⚙️ CAPL Compiler & Analyzer")
    with capl_reset_col:
        if st.button("🧼 Reset", key="reset_capl_selection"):
            st.session_state.selected_capl_file = "--Select CAPL file--"
            st.session_state.capl_last_analyzed_file = None
            st.session_state.capl_last_issues = None
            st.rerun()

    st.info(
        "Use sidebar selection to make CAPL source files available here. Then choose the CAPL file you want only in this tab.")
    show_current_sidebar_selection()
    show_help_popup('capl', [f for f in st.session_state.selected_files if f.lower().endswith((".can", ".txt"))])

    st.markdown("### Autonomous CAPL Agent Workspace")
    st.info("Use the CAPL agent system to run goal-driven workflows across uploaded documents, shared memory, and CAPL analysis.")
    agent_cols = st.columns(5)
    agent_defs = [
        ("Planning", "Breaks goals into steps"),
        ("Retrieval", "Queries FAISS memory"),
        ("Execution", "Runs tools/actions"),
        ("Reasoning", "Builds insights"),
        ("Coordination", "Merges final output"),
    ]
    for index, (agent_name, agent_role) in enumerate(agent_defs):
        with agent_cols[index]:
            st.markdown(f"**{agent_name} Agent**")
            st.caption(agent_role)
    st.text_input("Autonomous CAPL Goal", key="capl_autonomous_goal", placeholder="Analyze uploaded documents and give insights about risks")
    if st.button("🚀 Run Autonomous CAPL Agents", key="run_capl_agents"):
        with st.spinner("Running autonomous CAPL agents..."):
            st.session_state.capl_agent_result = run_capl_agent(
                st.session_state.capl_autonomous_goal,
                st.session_state.selected_files
            )
        st.success("Autonomous CAPL task completed.")

    if st.session_state.capl_agent_result:
        st.markdown("### Agent Output")
        st.markdown(st.session_state.capl_agent_result, unsafe_allow_html=True)

    if st.session_state.agent_run_history:
        with st.expander("Autonomous CAPL run history", expanded=False):
            for run in st.session_state.agent_run_history[-5:][::-1]:
                st.markdown(f"**{run['timestamp']}** — {html.escape(run['goal'])}")
                if run.get('plan'):
                    st.markdown(f"- Plan: {html.escape(', '.join(run['plan']))}")
                st.markdown("---")

    # Filter selected files for CAPL analysis
    capl_selectable_files = [f for f in st.session_state.selected_files if f.lower().endswith((".can", ".txt"))]
    active_capl_files = [] if st.session_state.selected_capl_file == "--Select CAPL file--" else [st.session_state.selected_capl_file]
    render_file_context_card("CAPL File Context", capl_selectable_files, active_capl_files)

    with st.expander("✍️ Create New CAPL Script", expanded=False):
        st.text_input("CAPL file name", key="capl_editor_name", help="Enter a file name ending with .can")
        st.text_area("Write CAPL code", key="capl_editor_code", height=260)
        live_editor_issues = analyze_capl_code_with_suggestions(st.session_state.capl_editor_code)

        st.markdown("### Live CAPL Preview")
        st.markdown(
            render_capl_code_with_highlights(st.session_state.capl_editor_code, live_editor_issues),
            unsafe_allow_html=True
        )

        st.markdown("### Live Issues & Suggestions")
        render_capl_issue_table(live_editor_issues)

        editor_ai_cols = st.columns([1, 4])
        with editor_ai_cols[0]:
            editor_ai_fix_clicked = st.button("🤖 Suggest Fix", key="capl_editor_ai_fix_button")

        if editor_ai_fix_clicked:
            llm = load_llm()
            # Add line numbers to code
            code_lines = st.session_state.capl_editor_code.split('\n')
            capl_code_with_line_numbers = '\n'.join(f"{i+1:4d}: {line}" for i, line in enumerate(code_lines))
            editor_prompt = f"""
    You are an advanced CAPL (CANoe/CANalyzer) AI IDE assistant.

    You act as:
    - CAPL static analyzer
    - CAPL debugger
    - CAPL refactoring engine
    - CAPL test generator
    - CAPL execution simulator
    - AI coding assistant (ChatGPT-like)

    ---

    ## 1. UNDERSTANDING TASK

    First analyze the CAPL code and understand:
    - purpose of the script
    - structure (events, timers, messages)
    - signals used
    - possible runtime flow

    ---

    ## 2. ISSUE DETECTION (MANDATORY)

    Detect:
    - syntax errors
    - logical errors
    - missing handlers
    - incorrect signal usage
    - bad practices

    STRICT RULE:
    - Every issue MUST include exact line number(s)
    - Do NOT guess lines

    ---

    ## 3. OUTPUT FORMAT (STRICT JSON)

    Return ONLY JSON:

    {{
      "chat_response": "Natural explanation of what the code does and key insights",

      "issues": [
    {{
      "title": "Short issue title",
      "description": "What is wrong and why",
      "severity": "error | warning | info",
      "line": 12
    }}
      ],

      "suggestions": [
    "Short action 1",
    "Short action 2",
    "Short action 3"
      ],

      "fixes": [
    {{
      "line": 12,
      "action": "replace | insert | delete",
      "new_code": "corrected CAPL snippet",
      "reason": "Why this fix is needed"
    }}
      ],

      "fixed_code": "FULL corrected CAPL script",

      "refactor": {{
    "improved_code": "cleaned and optimized CAPL code",
    "changes_summary": [
      "Improved naming",
      "Fixed structure",
      "Removed redundancy"
    ]
      }},

      "tests": [
    {{
      "name": "Test case name",
      "purpose": "What is being validated",
      "input": "Simulated input/event",
      "expected_output": "Expected behavior",
      "assertion": "Pass condition"
    }}
      ],

      "simulation": [
    {{
      "step": 1,
      "event": "CAPL event triggered",
      "action": "Function execution",
      "result": "State or signal change"
    }}
      ]
    }}

    ---

    ## 4. EXPLANATION FEATURE (IMPORTANT)

    For each issue, also support explanation:

    If user requests:
    👉 "Explain this error"

    Provide:
    - meaning
    - root cause
    - CAPL-specific reason
    - fix example

    ---

    ## 5. REFACTOR RULES

    When refactoring:
    - DO NOT change behavior
    - Improve structure only
    - Improve naming
    - Reduce duplication
    - Ensure CAPL compliance

    ---

    ## 6. FIX RULES

    When generating fixes:
    - minimal changes preferred
    - preserve intent
    - ensure compilable CAPL

    ---

    ## 7. SIMULATION RULES

    Simulate execution step-by-step:
    - events
    - timers
    - message handling
    - signal updates

    Return chronological flow.

    ---

    ## 8. SUGGESTIONS RULE

    Always include:
    - 3 to 6 short actionable next steps
    - must be context-aware
    - no generic suggestions

    ---

    ## 9. CHAT STYLE RULE

    chat_response must:
    - be natural (like ChatGPT)
    - explain what code does
    - highlight key behavior
    - be concise

    ---

    ## CAPL CODE INPUT:

    {capl_code_with_line_numbers}

    ---

    ## OPTIONAL CONTEXT:

    User Query:
    Suggest Fix

    Chat History:
    None

    Simulated Inputs:
    None
    """
            if llm is None:
                st.error("AI fix feature is unavailable because model backend could not be initialized.")
            else:
                with st.spinner("Generating CAPL fix suggestion..."):
                    try:
                        response = llm.invoke(editor_prompt)
                        # Parse JSON response
                        try:
                            ai_result = json.loads(response.strip())
                            st.session_state.capl_editor_ai_fix = ai_result.get("fixed_code", "")
                            st.session_state.capl_editor_ai_chat = ai_result.get("chat_response", "")
                            st.session_state.capl_editor_ai_issues = ai_result.get("issues", [])
                            st.session_state.capl_editor_ai_suggestions = ai_result.get("suggestions", [])
                        except json.JSONDecodeError:
                            st.error("AI response was not valid JSON. Using raw response as fix.")
                            st.session_state.capl_editor_ai_fix = response.strip()
                    except Exception as exc:
                        st.error(f"AI suggestion failed: {exc}")
                        st.session_state.capl_editor_ai_fix = ""

        if st.session_state.capl_editor_ai_fix:
            if st.session_state.get("capl_editor_ai_chat"):
                st.markdown("### 🤖 AI Analysis")
                st.markdown(st.session_state.capl_editor_ai_chat)
        
            if st.session_state.get("capl_editor_ai_issues"):
                st.markdown("### Issues Detected")
                for issue in st.session_state.capl_editor_ai_issues:
                    severity_icon = {"error": "❌", "warning": "⚠️", "info": "ℹ️"}.get(issue.get("severity", "info"), "ℹ️")
                    st.markdown(f"{severity_icon} **{issue.get('title', 'Issue')}** (Line {issue.get('line', '?')}): {issue.get('description', '')}")
        
            if st.session_state.get("capl_editor_ai_suggestions"):
                st.markdown("### Suggestions")
                for sugg in st.session_state.capl_editor_ai_suggestions:
                    st.markdown(f"- {sugg}")
        
            st.markdown("### Suggested Corrected CAPL Code")
            fixed_editor_issues = analyze_capl_code_with_suggestions(st.session_state.capl_editor_ai_fix)
            st.markdown(
                render_capl_code_with_highlights(st.session_state.capl_editor_ai_fix, fixed_editor_issues),
                unsafe_allow_html=True
            )
            if st.button("Use Suggested Fix", key="capl_editor_use_ai_fix"):
                st.session_state.capl_editor_code = st.session_state.capl_editor_ai_fix
                st.rerun()

        if st.button("💾 Save New CAPL Script"):
            new_file_name = st.session_state.capl_editor_name.strip()
            if not new_file_name:
                st.warning("Enter a file name for the CAPL script.")
            else:
                if not new_file_name.lower().endswith(".can"):
                    new_file_name += ".can"

                file_bytes = st.session_state.capl_editor_code.encode("utf-8")
                existing_index = next(
                    (idx for idx, file_info in enumerate(st.session_state.uploaded_files) if
                     file_info["name"] == new_file_name),
                    None
                )

                if existing_index is None:
                    st.session_state.uploaded_files.append({"name": new_file_name, "bytes": file_bytes})
                else:
                    st.session_state.uploaded_files[existing_index] = {"name": new_file_name, "bytes": file_bytes}

                st.session_state.file_texts[new_file_name] = st.session_state.capl_editor_code
                if new_file_name not in st.session_state.selected_files:
                    st.session_state.selected_files.append(new_file_name)
                st.session_state.capl_last_analyzed_file = None
                st.session_state.capl_last_issues = None
                st.session_state.capl_editor_ai_fix = ""
                st.success(f"Saved {new_file_name} and added it to the selected files.")

    use_all_txt = st.checkbox("Include all .txt files as CAPL")
    with st.spinner("Loading CAPL files..."):
        ensure_files_processed(capl_selectable_files)

    if use_all_txt:
        capl_files = [
            f for f in capl_selectable_files
            if f.lower().endswith((".can", ".txt"))
        ]
    else:
        capl_files = [
            f for f in capl_selectable_files
            if f.lower().endswith((".can", ".txt")) and
               is_capl_code(st.session_state.file_texts.get(f, ""))
        ]

    if not capl_files:
        st.warning("Upload/select CAPL (.can/.capl) files")
    else:
        capl_options = ["--Select CAPL file--"] + capl_files
        if st.session_state.selected_capl_file not in capl_options:
            st.session_state.selected_capl_file = "--Select CAPL file--"
        selected_capl = st.selectbox("Select CAPL file", capl_options, key="selected_capl_file")
        capl_selected = selected_capl != "--Select CAPL file--"
        action_cols = st.columns([1, 1, 4])
        with action_cols[0]:
            analyze_clicked = st.button("🚀 Compile & Analyze", disabled=not capl_selected)
        with action_cols[1]:
            clear_analysis = st.button("🧹 Clear Analysis", disabled=not capl_selected)

        if not capl_selected:
            st.info("Choose a CAPL file to view and analyze.")
        else:
            if clear_analysis:
                st.session_state.capl_last_analyzed_file = None
                st.session_state.capl_last_issues = None

            code = st.session_state.file_texts[selected_capl]
            issues = []

            if analyze_clicked:
                issues = analyze_capl_code_with_suggestions(code)
                st.session_state.capl_last_analyzed_file = selected_capl
                st.session_state.capl_last_issues = issues

            if st.session_state.capl_last_analyzed_file == selected_capl and st.session_state.capl_last_issues is not None:
                issues = st.session_state.capl_last_issues

            st.markdown("### 📄 CAPL Code")
            st.markdown(render_capl_code_with_highlights(code, issues), unsafe_allow_html=True)

            if st.session_state.capl_last_analyzed_file == selected_capl and st.session_state.capl_last_issues is not None:
                st.markdown("### 🛠️ Issues Found")
                render_capl_issue_table(issues)

                # 🤖 AI Suggestions
                if st.checkbox("🤖 AI Suggestions"):
                    llm = load_llm()
                    if llm is None:
                        st.error("AI fix feature is unavailable because model backend could not be initialized.")
                    else:
                        # Add line numbers to code
                        code_lines = code.split('\n')
                        capl_code_with_line_numbers = '\n'.join(f"{i+1:4d}: {line}" for i, line in enumerate(code_lines))
                        prompt = f"""
    You are an advanced CAPL (CANoe/CANalyzer) AI IDE assistant.

    You act as:
    - CAPL static analyzer
    - CAPL debugger
    - CAPL refactoring engine
    - CAPL test generator
    - CAPL execution simulator
    - AI coding assistant (ChatGPT-like)

    ---

    ## 1. UNDERSTANDING TASK

    First analyze the CAPL code and understand:
    - purpose of the script
    - structure (events, timers, messages)
    - signals used
    - possible runtime flow

    ---

    ## 2. ISSUE DETECTION (MANDATORY)

    Detect:
    - syntax errors
    - logical errors
    - missing handlers
    - incorrect signal usage
    - bad practices

    STRICT RULE:
    - Every issue MUST include exact line number(s)
    - Do NOT guess lines

    ---

    ## 3. OUTPUT FORMAT (STRICT JSON)

    Return ONLY JSON:

    {{
      "chat_response": "Natural explanation of what the code does and key insights",

      "issues": [
    {{
      "title": "Short issue title",
      "description": "What is wrong and why",
      "severity": "error | warning | info",
      "line": 12
    }}
      ],

      "suggestions": [
    "Short action 1",
    "Short action 2",
    "Short action 3"
      ],

      "fixes": [
    {{
      "line": 12,
      "action": "replace | insert | delete",
      "new_code": "corrected CAPL snippet",
      "reason": "Why this fix is needed"
    }}
      ],

      "fixed_code": "FULL corrected CAPL script",

      "refactor": {{
    "improved_code": "cleaned and optimized CAPL code",
    "changes_summary": [
      "Improved naming",
      "Fixed structure",
      "Removed redundancy"
    ]
      }},

      "tests": [
    {{
      "name": "Test case name",
      "purpose": "What is being validated",
      "input": "Simulated input/event",
      "expected_output": "Expected behavior",
      "assertion": "Pass condition"
    }}
      ],

      "simulation": [
    {{
      "step": 1,
      "event": "CAPL event triggered",
      "action": "Function execution",
      "result": "State or signal change"
    }}
      ]
    }}

    ---

    ## 4. EXPLANATION FEATURE (IMPORTANT)

    For each issue, also support explanation:

    If user requests:
    👉 "Explain this error"

    Provide:
    - meaning
    - root cause
    - CAPL-specific reason
    - fix example

    ---

    ## 5. REFACTOR RULES

    When refactoring:
    - DO NOT change behavior
    - Improve structure only
    - Improve naming
    - Reduce duplication
    - Ensure CAPL compliance

    ---

    ## 6. FIX RULES

    When generating fixes:
    - minimal changes preferred
    - preserve intent
    - ensure compilable CAPL

    ---

    ## 7. SIMULATION RULES

    Simulate execution step-by-step:
    - events
    - timers
    - message handling
    - signal updates

    Return chronological flow.

    ---

    ## 8. SUGGESTIONS RULE

    Always include:
    - 3 to 6 short actionable next steps
    - must be context-aware
    - no generic suggestions

    ---

    ## 9. CHAT STYLE RULE

    chat_response must:
    - be natural (like ChatGPT)
    - explain what code does
    - highlight key behavior
    - be concise

    ---

    ## CAPL CODE INPUT:

    {capl_code_with_line_numbers}

    ---

    ## OPTIONAL CONTEXT:

    User Query:
    Analyze and fix CAPL code

    Chat History:
    None

    Simulated Inputs:
    None
    """
                        with st.spinner("Analyzing and fixing with AI..."):
                            try:
                                response = llm.invoke(prompt)
                                # Parse JSON response
                                try:
                                    ai_result = json.loads(response.strip())
                                    corrected_code = ai_result.get("fixed_code", "")
                                    if corrected_code:
                                        # Update the code in session state
                                        st.session_state.file_texts[selected_capl] = corrected_code
                                        code = corrected_code
                                        issues = analyze_capl_code_with_suggestions(code)
                                        st.session_state.capl_last_analyzed_file = selected_capl
                                        st.session_state.capl_last_issues = issues
                                        st.success("✅ Code corrected by AI!")
                                        st.markdown("### 🤖 AI Analysis")
                                        st.markdown(ai_result.get("chat_response", ""))
                                        if ai_result.get("issues"):
                                            st.markdown("### Issues Detected")
                                            for issue in ai_result["issues"]:
                                                severity_icon = {"error": "❌", "warning": "⚠️", "info": "ℹ️"}.get(issue.get("severity", "info"), "ℹ️")
                                                st.markdown(f"{severity_icon} **{issue.get('title', 'Issue')}** (Line {issue.get('line', '?')}): {issue.get('description', '')}")
                                        if ai_result.get("suggestions"):
                                            st.markdown("### Suggestions")
                                            for sugg in ai_result["suggestions"]:
                                                st.markdown(f"- {sugg}")
                                        st.markdown("### 📄 Corrected CAPL Code")
                                        st.markdown(render_capl_code_with_highlights(code, issues), unsafe_allow_html=True)
                                        if issues:
                                            st.warning("⚠️ Some issues remain:")
                                        render_capl_issue_table(issues)
                                except json.JSONDecodeError:
                                    st.error("AI response was not valid JSON.")
                            except Exception as exc:
                                st.error(f"AI suggestion failed: {exc}")

    st.markdown('</div>', unsafe_allow_html=True)
