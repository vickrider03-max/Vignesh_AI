README.txt
==========

🚀 Multi-Utility File & CAPL Analyzer Tool
=========================================

Overview
--------
This Streamlit-based application helps manage, analyze, and compare files, with special support for CAPL scripts. It combines file dashboards, comparisons, CAPL analysis, and AI-assisted code fixing in a single platform.

---

App Layout & Workflow
--------------------

+----------------------+
| Sidebar              |
|----------------------|
| Upload Files         |
| Select Files         |
| Filter CAPL (.can)   |
+----------------------+
          |
          v
+------------------------------------------+
| Main Tabs:                                |
|------------------------------------------|
| [ Chat ]   [ Dashboard ]   [ Compare ]   [ CAPL ] |
+------------------------------------------+

Tab Workflow:
-------------

1. Chat Tab
   +-----------------------------+
   | Ask questions about files   |
   | Semantic / AI answers       |
   +-----------------------------+
          |
          v
   (Uses uploaded files & AI backend)

2. Dashboard Tab
   +------------------------------+
   | Select a file from sidebar   |
   | Visualize trends (Excel/CSV)|
   | Aggregated stats             |
   +------------------------------+
          |
          v
   (Optional download of analyzed results)

3. Compare Tab
   +--------------------------------+
   | Multi-file selection           |
   | Inline word-level differences  |
   | Download comparison Excel      |
   +--------------------------------+
          |
          v
   (At least 2 files required)

4. CAPL Tab
   +-------------------------------------------+
   | Select existing CAPL file or create new  |
   | Compile & analyze code                     |
   | View issues / suggestions                 |
   | AI-assisted fix / Apply fix / Save file   |
   +-------------------------------------------+
          |
          v
   (Updates session state & file texts)

---

Feature Summary
---------------

1. **Chat / RAG Interface**
   - Ask questions about selected files.
   - Context-aware AI responses.

2. **File Dashboard**
   - Test report analysis and visualization.
   - Downloadable Excel summaries.

3. **Compare Files**
   - Multi-file comparison.
   - Inline word-level differences.
   - Downloadable Excel comparison.

4. **CAPL Compiler & Analyzer**
   - Upload or create CAPL scripts (.can/.txt).
   - Syntax highlighting & code analysis.
   - AI-assisted code fixes.
   - Save new or corrected CAPL files.

5. **Interactive UI**
   - Tabs for workflows.
   - Reset buttons for selections and results.
   - Expandable live editor for CAPL scripts.

6. **AI Integration**
   - Auto-correct CAPL code.
   - Chat-based file analysis.

7. **Session Management**
   - Tracks uploaded files and selected files per tab.
   - Maintains last analyzed CAPL file and issues.

---

How to Use
----------

1. **Setup**
   - Python >= 3.10
   - Install dependencies:
     ```
     pip install streamlit openai pandas plotly
     ```
   - Configure AI backend / API keys if using AI features.

2. **Run**
   - streamlit run app.py
   
3. **Sidebar**
- Upload files.
- Select files to be available in tabs.
- Optionally filter CAPL scripts.

4. **Tabs**
- **Chat** – Ask questions about uploaded files.
- **Dashboard** – Visualize file content, trends, and statistics.
- **Compare** – Choose 2+ files and see word-level differences.
- **CAPL** – Edit, compile, analyze CAPL scripts, AI fixes, save.

5. **CAPL AI Fix**
- Click "🤖 Suggest Fix" → review AI suggestion → click "Use Suggested Fix".

6. **Reset Buttons**
- Clear selections and results in each tab.

---

Tips & Notes
------------

- Limit comparisons to <2000 lines for performance.
- CAPL files must end with `.can`.
- AI features require backend availability.
- Use “Include all .txt files as CAPL” cautiously.

---

ASCII Workflow Example
---------------------

Sidebar:
--------
+----------------------+
| Upload Files         |
| Select Files         |
| Filter CAPL (.can)   |
+----------------------+

Tabs:
-----
[ Chat ] -> AI Q&A using uploaded files
[ Dashboard ] -> File stats & trends -> Excel download
[ Compare ] -> Multi-file diff -> HTML + Excel
[ CAPL ] -> Edit/Compile/Analyze -> AI Fix -> Save

CAPL AI Flow:
-------------
Create/Edit CAPL
    |
    v
Compile & Analyze
    |
    v
AI Suggest Fix? -- Yes --> Review & Apply --> Update Editor
    |
    No
    v
Save CAPL Script

---

Credits
-------
- Built with Streamlit, Pandas, Plotly, OpenAI API
- CAPL analysis inspired by automotive testing
- AI auto-fix powered by LLMs

---

Contact
-------
For support or feedback, contact the project maintainer.