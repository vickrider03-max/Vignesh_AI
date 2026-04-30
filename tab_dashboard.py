# Auto-generated from legacy_app.py during modular refactor.
# The original monolith is retained as rollback documentation.

from functions import *

# ==============================
# DASHBOARD TAB UI
# Moved from legacy_app.py tab body.
# UI rendering and event handling live here; backend work is delegated to functions.py.
# ==============================
def render_dashboard_tab():
    st.markdown('<div id="dashboard-section">', unsafe_allow_html=True)
    dashboard_header_col, dashboard_reset_col = st.columns([8, 1])
    with dashboard_header_col:
        st.subheader("Dashboard")
    with dashboard_reset_col:
        if st.button("🧼 Reset", key="reset_dashboard_selection"):
            st.session_state.file_dropdown = "--Select File--"
            st.session_state.dashboard_chart_type = "Pie Chart"
            st.session_state.dashboard_bar_orientation = "Vertical"
            st.rerun()

    show_current_sidebar_selection()
    show_help_popup('dashboard', [
        f for f in st.session_state.selected_files
        if f.lower().endswith((".html", ".htm", ".xlsx"))
    ])

    # Filter selected files for dashboard-compatible formats
    dashboard_files = [
        f for f in st.session_state.selected_files
        if f.lower().endswith((".html", ".htm", ".xlsx"))
    ]
    active_dashboard_files = [] if st.session_state.file_dropdown == "--Select File--" else [st.session_state.file_dropdown]
    render_file_context_card("Dashboard File Context", dashboard_files, active_dashboard_files)


    def clean_text(x):
        return re.sub(r"\s+", " ", x).strip().lower()


    def safe_int(x):
        nums = re.findall(r"\d+", str(x))
        return int(nums[0]) if nums else 0


    def sort_key(x):
        return [int(i) for i in x.split(".")]


    def format_label(x):
        level = x.count(".")
        return "   " * level + x


    def extract_timestamp_from_line(line):
        """Extract timestamp or identifier from a verdict line"""
        # Try to extract decimal number first (e.g., 35.877473)
        decimal_match = re.search(r'\b(\d+\.\d+)\b', line)
        if decimal_match:
            return decimal_match.group(1)

        # Try to extract time pattern (HH:MM:SS or HH:MM)
        time_match = re.search(r'\b(\d{1,2}:\d{2}:\d{2})\b', line)
        if time_match:
            return time_match.group(1)

        # Try to extract date-time pattern
        datetime_match = re.search(r'\d{4}-\d{2}-\d{2}\s+\d{1,2}:\d{2}:\d{2}', line)
        if datetime_match:
            return datetime_match.group(0)

        # Try to extract test case ID (e.g., "1.", "Test Case 1", "TC_001")
        id_match = re.search(r'(?:test\s+case|tc)[\s_-]*(\d+)', line, re.IGNORECASE)
        if id_match:
            return f"TC {id_match.group(1)}"

        # Extract first number or identifier
        num_match = re.search(r'^\d+', line)
        if num_match:
            return num_match.group(0)

        # Extract the first meaningful text before colon or special char
        first_word = re.match(r'^([a-zA-Z0-9_\-\.]+)', line)
        if first_word:
            return first_word.group(1)

        # Fallback to full line (truncated)
        return line[:30]


    def extract_test_results_grouped(soup):
        """
        Extract test results from HTML grouped by test fixtures.
        Dynamically discovers fixtures from HTML structure and their numerical counts.
        Only counts the main test case verdict, not sub-step verdicts.
        Captures detailed failure information for failed test cases (step ID, action, timestamp).
        """
        results = {}

        # First, extract fixture names and counts from GroupHeadingTable structures
        group_tables = soup.find_all('table', class_='GroupHeadingTable')

        for group_table in group_tables:
            try:
                rows = group_table.find_all('tr')
                if len(rows) >= 2:
                    # First row contains fixture name in Heading3
                    first_row = rows[0]
                    heading = first_row.find('big', class_='Heading3')

                    if heading:
                        heading_text = heading.get_text(strip=True)
                        fixture_match = re.search(r'Test Fixture:\s*(.+?)(?:\s|$)', heading_text, re.IGNORECASE)

                        if fixture_match:
                            fixture_name = fixture_match.group(1).strip()

                            # Second row contains the count in OverviewResultTable
                            second_row = rows[1]
                            overview_table = second_row.find('table', class_='OverviewResultTable')

                            if overview_table:
                                count_cell = overview_table.find('td')
                                if count_cell:
                                    try:
                                        count = int(count_cell.get_text(strip=True))

                                        if fixture_name not in results:
                                            results[fixture_name] = {
                                                "name": fixture_name,
                                                "test_cases": [],
                                                "pass": count,  # All are passed by default from green cells
                                                "fail": 0,
                                                "error": 0,
                                                "not executed": 0,
                                                "inconclusive": 0,
                                                "total": count,
                                                "count_cell_class": count_cell.get('class', [''])[0]
                                            }
                                    except ValueError:
                                        pass
            except Exception as e:
                pass

        # Parse full text for verdict distribution
        full_text = soup.get_text("\n", strip=True)
        lines = [l.strip() for l in full_text.split("\n") if l.strip()]

        current_fixture = None

        for i, line in enumerate(lines):
            line_lower = line.lower()

            # Check for fixture marker
            if "test fixture:" in line_lower:
                fixture_match = re.search(r'Test Fixture:\s*(.+?)(?:\s|$)', line, re.IGNORECASE)
                if fixture_match:
                    current_fixture = fixture_match.group(1).strip()
                    if current_fixture not in results:
                        results[current_fixture] = {
                            "name": current_fixture,
                            "test_cases": [],
                            "pass": 0,
                            "fail": 0,
                            "error": 0,
                            "not executed": 0,
                            "inconclusive": 0,
                            "total": 0
                        }

            # Check for test case number with verdict on SAME line (e.g., "1.2.1 ...description...: Passed")
            # Only match if line contains both test number and a verdict
            elif re.match(r'^\d+\.\d+', line) and current_fixture:
                verdict_match = re.search(r':\s*(Passed|Failed|Pass|Fail|Error|Not Executed|Inconclusive)\s*$', line,
                                          re.IGNORECASE)

                if verdict_match:
                    verdict_word = verdict_match.group(1).lower()

                    # Normalize verdict and increment counter
                    if "pass" in verdict_word:
                        verdict_type = "Pass"
                        results[current_fixture]["pass"] += 1
                    elif "fail" in verdict_word:
                        verdict_type = "Failed"
                        results[current_fixture]["fail"] += 1
                    elif "error" in verdict_word:
                        verdict_type = "Error"
                        results[current_fixture]["error"] += 1
                    elif "not executed" in verdict_word:
                        verdict_type = "Not Executed"
                        results[current_fixture]["not executed"] += 1
                    elif "inconclusive" in verdict_word:
                        verdict_type = "Inconclusive"
                        results[current_fixture]["inconclusive"] += 1
                    else:
                        continue

                    # Now look forward for timestamp and failure details
                    timestamp = None
                    test_step = "Step"
                    failure_step_id = ""

                    def score_timestamp(candidate):
                        if not candidate:
                            return -1
                        parts = candidate.split('.')
                        if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
                            return -1
                        leading_num = int(parts[0])
                        decimal_places = len(parts[1])
                        decimal_bonus = 10000 if decimal_places >= 3 else (100 if decimal_places == 2 else 0)
                        return decimal_bonus + leading_num

                    def find_best_timestamp(text):
                        # Find all decimal numbers
                        matches = re.findall(r'\b(\d+\.\d+)\b', text)
                        if not matches:
                            return None

                        return max(matches, key=score_timestamp)

                    def find_first_relevant_timestamp(text):
                        # Prefer first high-precision timestamp (3+ decimals) in text
                        for m in re.findall(r'\b(\d+\.\d+)\b', text):
                            if len(m.split('.')[1]) >= 3:
                                return m
                        # fallback to lower precision (2+ decimals)
                        for m in re.findall(r'\b(\d+\.\d+)\b', text):
                            if len(m.split('.')[1]) >= 2:
                                return m
                        return None

                    def consider_timestamp(candidate):
                        nonlocal timestamp
                        if not candidate:
                            return
                        # Prefer the first nearby relevant timestamp (prefer earlier block context)
                        if not timestamp:
                            timestamp = candidate
                            return
                        # if already have a high precision timestamp, keep it
                        if len(timestamp.split('.')[1]) >= 3:
                            return
                        # if candidate is higher precision than existing timestamp, replace
                        if len(candidate.split('.')[1]) > len(timestamp.split('.')[1]):
                            timestamp = candidate
                            return
                        # as fallback, use original scoring strategy
                        if score_timestamp(candidate) > score_timestamp(timestamp):
                            timestamp = candidate

                    same_line_step = re.search(r'(\d+(?:\.\d+)+)\.\s+([^:]+):\s*(failed|fail|error)', line,
                                               re.IGNORECASE)
                    if same_line_step:
                        failure_step_id = same_line_step.group(1)
                        action_text = same_line_step.group(2).strip()
                        test_step = action_text  # prefer just action text as details
                        # timestamp candidate in same line: prefer first relevant and then best
                        consider_timestamp(find_first_relevant_timestamp(line) or find_best_timestamp(line))

                    for k in range(i + 1, min(i + 150, len(lines))):
                        next_line = lines[k]

                        # Stop if we hit next test case
                        if re.match(r'^\d+\.\d+(?:\s|$)', next_line) and k > i + 5:
                            break

                        # Look for timestamp (decimal number) - general case
                        consider_timestamp(find_first_relevant_timestamp(next_line) or find_best_timestamp(next_line))

                        # For Failed/Error verdicts, look for failed step details
                        if verdict_type in ["Failed", "Error"] and not failure_step_id:
                            next_line_lower = next_line.lower()

                            # Look for step identifier with action (e.g., "10.1.6.9.4. Await Value Match: Failed")
                            step_match = re.search(r'(\d+(?:\.\d+)+)\.\s+([^:]+):\s*(failed|fail|error)', next_line,
                                                   re.IGNORECASE)
                            if step_match:
                                failure_step_id = step_match.group(1)
                                action_text = step_match.group(2).strip()
                                test_step = action_text  # details should not carry step ID prefix
                                # Extract timestamp from the failure step line if present using best candidate selection
                                consider_timestamp(find_best_timestamp(next_line))
                            else:
                                # Look for common failure indicators in other formats
                                if any(keyword in next_line_lower for keyword in
                                       ["condition", "value", "expected", "actual", "mismatch", "not found",
                                        "exception", "error", "failed to", "failed"]):
                                    if not re.match(r'^\d+\.\d+', next_line):
                                        # Extract step number if it starts with numbers
                                        step_num_match = re.match(r'^(\d+(?:\.\d+)*)', next_line.strip())
                                        if step_num_match:
                                            failure_step_id = step_num_match.group(1)
                                            test_step = next_line[:80]

                        # Look for action keywords for passed cases
                        if verdict_type == "Pass":
                            next_line_lower = next_line.lower()

                            if "execute" in next_line_lower:
                                match = re.search(r'execute\s+(\w+)', next_line_lower)
                                if match:
                                    test_step = match.group(1).capitalize()
                            elif "question" in next_line_lower and "text" in next_line_lower:
                                test_step = "Question/Text"
                            elif "await" in next_line_lower or "wait" in next_line_lower:
                                test_step = "Await Value Match"
                            elif "resume" in next_line_lower:
                                test_step = "Resume"
                            elif "set" in next_line_lower:
                                test_step = "Set"
                            elif "tester" in next_line_lower and "confirmed" in next_line_lower:
                                test_step = "Tester Confirmation"

                    # Add test case if we found a timestamp
                    if timestamp:
                        results[current_fixture]["test_cases"].append({
                            "timestamp": timestamp,
                            "verdict": verdict_type,
                            "details": test_step
                        })

        # Calculate totals
        for fixture_name in results:
            # Use the maximum of: parsed test cases OR initial summary count
            # This preserves the fixture summary count if detailed parsing didn't find individual cases
            parsed_count = len(results[fixture_name]["test_cases"])
            initial_count = results[fixture_name].get("total", 0)
            results[fixture_name]["total"] = max(parsed_count, initial_count)

        return results


    if not st.session_state.selected_files:
        st.info("Select files from the sidebar to show dashboard.")
    elif not dashboard_files:
        st.info("No dashboard-friendly files selected. Choose HTML/HTM/XLSX in sidebar for dashboard details.")
    else:
        st.info(
            "Files selected in the sidebar are available here. Choose only the dashboard file you want to inspect in this tab.")
        dashboard_options = ["--Select File--"] + dashboard_files
        if st.session_state.file_dropdown not in dashboard_options:
            st.session_state.file_dropdown = "--Select File--"
        file_dropdown = st.selectbox("Select a dashboard file", dashboard_options, key="file_dropdown")

        if file_dropdown != "--Select File--":
            with st.spinner("Loading dashboard file..."):
                ensure_file_processed(file_dropdown)
            file_entry = get_uploaded_file_entry(file_dropdown)
            file_bytes = file_entry["bytes"] if file_entry else b""
            chart_type = st.radio(
                "Chart type",
                ["Pie Chart", "Bar Chart"],
                index=0,
                horizontal=True,
                key="dashboard_chart_type",
            )
            bar_orientation = "Vertical"
            if chart_type == "Bar Chart":
                bar_orientation = st.radio(
                    "Bar orientation",
                    ["Vertical", "Horizontal"],
                    index=0,
                    horizontal=True,
                    key="dashboard_bar_orientation",
                )
            horizontal_bar = (bar_orientation == "Horizontal")

            if file_dropdown.lower().endswith(".xlsx"):
                data = st.session_state.excel_data_by_file.get(file_dropdown, [])

                if data:
                    columns = ["--Select Column--"] + list(data[0].keys())
                    col = st.selectbox("Column to analyze", columns, index=0)

                    if col != "--Select Column--":
                        counts = get_column_counts(data, col)
                        if counts:
                            fig = plot_pie_chart(counts,
                                                 f"{col} Distribution") if chart_type == "Pie Chart" else plot_bar_chart(
                                counts, f"{col} Distribution", horizontal=horizontal_bar
                            )
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.warning("No data found for the selected column.")
                else:
                    st.warning("No Excel data available for analysis.")

            elif file_dropdown.lower().endswith((".html", ".htm")):
                st.markdown("### 🔐 Login Info")
                login = extract_login_name_from_html(file_bytes)
                st.write("Login Name:", login)

                st.markdown("### 📊 Statistics")
                stats = extract_statistics_from_html(file_bytes)

                if any(v > 0 for v in stats.values()):
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Executed", stats["Executed"])
                    c2.metric("Passed", stats["Passed"])
                    c3.metric("Failed", stats["Failed"])

                    fig = plot_pie_chart(stats, "Statistics") if chart_type == "Pie Chart" else plot_bar_chart(
                        stats, "Statistics", horizontal=horizontal_bar
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("⚠️ Could not extract statistics")

                st.markdown("### 📋 Test Results")
                grouped_results = extract_test_results_grouped_from_html(file_bytes)

                if grouped_results:
                    st.markdown("#### 👉 Executed Test Cases Summary")
                    all_fixtures = sorted(grouped_results.keys())
                    max_cols = 5
                    num_cols = min(len(all_fixtures), max_cols)

                    if num_cols > 0:
                        cols = st.columns(num_cols)
                        for idx, fixture_name in enumerate(all_fixtures[:max_cols]):
                            if fixture_name in grouped_results:
                                total_cases = grouped_results[fixture_name].get("total", 0)
                                pass_count = grouped_results[fixture_name].get("pass", 0)
                                with cols[idx % num_cols]:
                                    st.metric(
                                        label=fixture_name,
                                        value=total_cases,
                                        delta=f"{pass_count} Passed" if total_cases > 0 else None,
                                        delta_color="inverse"
                                    )

                        if len(all_fixtures) > max_cols:
                            st.markdown(f"**And {len(all_fixtures) - max_cols} more fixtures...**")

                    st.markdown("---")
                    st.markdown("#### Full Test Fixtures Overview")

                    fixture_data = []
                    for fixture_name in sorted(grouped_results.keys()):
                        data = grouped_results[fixture_name]
                        fixture_data.append({
                            "Test Fixture": fixture_name,
                            "✅ Pass": data.get("pass", 0),
                            "❌ Fail": data.get("fail", 0),
                            "⚠️ Error": data.get("error", 0),
                            "⏭️ Not Executed": data.get("not executed", 0),
                            "❓ Inconclusive": data.get("inconclusive", 0),
                            "📊 Total": data.get("total", 0)
                        })

                    df = pd.DataFrame(fixture_data)


                    def style_fixture_table(row):
                        styles = [''] * len(row)
                        if '✅ Pass' in df.columns:
                            pass_col_idx = df.columns.get_loc('✅ Pass')
                            if row['✅ Pass'] > 0:
                                styles[pass_col_idx] = 'background-color: #90EE90; color: black; font-weight: bold;'
                        if '❌ Fail' in df.columns:
                            fail_col_idx = df.columns.get_loc('❌ Fail')
                            if row['❌ Fail'] > 0:
                                styles[fail_col_idx] = 'background-color: #FF6B6B; color: white; font-weight: bold;'
                        if '⚠️ Error' in df.columns:
                            error_col_idx = df.columns.get_loc('⚠️ Error')
                            if row['⚠️ Error'] > 0:
                                styles[error_col_idx] = 'background-color: #FF8C8C; color: white; font-weight: bold;'
                        if '⏭️ Not Executed' in df.columns:
                            notexec_col_idx = df.columns.get_loc('⏭️ Not Executed')
                            if row['⏭️ Not Executed'] > 0:
                                styles[notexec_col_idx] = 'background-color: #FFD700; color: black;'
                        if '❓ Inconclusive' in df.columns:
                            inconc_col_idx = df.columns.get_loc('❓ Inconclusive')
                            if row['❓ Inconclusive'] > 0:
                                styles[inconc_col_idx] = 'background-color: #FFA500; color: black;'
                        return styles


                    styled_fixture_df = df.style.apply(style_fixture_table, axis=1)
                    st.dataframe(styled_fixture_df, use_container_width=True)

                    st.markdown("#### View Test Cases by Fixture")
                    fixture_options = sorted(grouped_results.keys())
                    selected_fixture = st.selectbox("Select Test Fixture:", fixture_options, key="fixture_select")

                    if selected_fixture:
                        fixture_info = grouped_results[selected_fixture]
                        st.markdown(f"### 📋 Test Fixture: **{selected_fixture}**")

                        col1, col2, col3, col4, col5 = st.columns(5)
                        col1.metric("✅ Pass", fixture_info.get("pass", 0))
                        col2.metric("❌ Fail", fixture_info.get("fail", 0))
                        col3.metric("⚠️ Error", fixture_info.get("error", 0))
                        col4.metric("⏭️ Not Executed", fixture_info.get("not executed", 0))
                        col5.metric("❓ Inconclusive", fixture_info.get("inconclusive", 0))

                        mode = st.radio("Show Test Cases", ["All", "Passed only", "Failed/Error only"], index=2,
                                        key="test_case_mode")

                        if fixture_info["test_cases"]:
                            test_cases_df = pd.DataFrame()
                            if mode == "All":
                                test_cases_to_show = fixture_info["test_cases"]
                                heading = "Test Cases (All)"
                            elif mode == "Passed only":
                                test_cases_to_show = [
                                    tc for tc in fixture_info["test_cases"]
                                    if tc.get("verdict", "").lower() in ["pass", "passed"]
                                ]
                                heading = "Test Cases (Passed Only)"
                            else:
                                test_cases_to_show = [
                                    tc for tc in fixture_info["test_cases"]
                                    if tc.get("verdict", "").lower() in ["failed", "fail", "error"]
                                ]
                                heading = "Test Cases (Failed/Error Only)"

                            st.markdown(f"#### {heading}")
                            if not test_cases_to_show:
                                st.info("No test cases to display for selected filter.")
                            else:
                                test_cases_df = pd.DataFrame(test_cases_to_show)


                            def color_verdict(val):
                                if val == "Pass":
                                    return "background-color: #90EE90; color: black; font-weight: bold;"
                                if val in ["Fail", "Failed", "Error"]:
                                    return "background-color: #FF6B6B; color: white; font-weight: bold;"
                                if val == "Not Executed":
                                    return "background-color: #FFD700; color: black;"
                                if val == "Inconclusive":
                                    return "background-color: #FFA500; color: black;"
                                return ""


                            if not test_cases_df.empty and "verdict" in test_cases_df.columns:
                                styled_df = test_cases_df.style.map(
                                    lambda x: color_verdict(x) if isinstance(x, str) else "",
                                    subset=["verdict"]
                                )
                                st.dataframe(styled_df, use_container_width=True, hide_index=True)
                            elif not test_cases_df.empty:
                                st.dataframe(test_cases_df, use_container_width=True, hide_index=True)

                        verdict_counts = {
                            "Pass": fixture_info.get("pass", 0),
                            "Fail": fixture_info.get("fail", 0),
                            "Error": fixture_info.get("error", 0),
                            "Not Executed": fixture_info.get("not executed", 0),
                            "Inconclusive": fixture_info.get("inconclusive", 0)
                        }
                        verdict_counts = {k: v for k, v in verdict_counts.items() if v > 0}

                        if verdict_counts:
                            fig = plot_pie_chart(verdict_counts,
                                                 f"Verdict Distribution - {selected_fixture}") if chart_type == "Pie Chart" else plot_bar_chart(
                                verdict_counts, f"Verdict Distribution - {selected_fixture}", horizontal=horizontal_bar
                            )
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("No structured test results found.")

    st.markdown('</div>', unsafe_allow_html=True)

    # -------------------------------
