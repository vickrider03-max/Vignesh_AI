# IntelliDoc AI Control Room

IntelliDoc AI is a Streamlit-based document intelligence workspace for uploading,
previewing, searching, chatting with, comparing, and analyzing documents. It also
includes a dedicated CAPL workspace for CANoe/CAPL script review, issue detection,
AI-assisted fixes, and autonomous CAPL task runs.

The current app is implemented in `app.py`, with `Backup/legacy_app.py` retained as
rollback documentation from the original monolithic version.

---

## Core Features

### Authentication and Session Workspace
- Login/logout flow with user and creator roles.
- Usage timer, role display, available-file count, and session status strip.
- Creator-only login/logout history and active-user statistics.
- Logout cleanup for uploaded files, chat history, workspace memory, caches, and preview data.
- Mobile workspace toggle for opening uploads/files on small screens.

### Sidebar File Management
- Multi-file upload from the sidebar.
- File cards with selection state, preview links, delete actions, and clear-all support.
- Selected files are shared across Chat, Dashboard, Compare, and CAPL tabs.
- SHA256-based change detection and in-memory LRU caches reduce repeat processing.

Supported upload types:

`pdf`, `doc`, `docx`, `txt`, `md`, `log`, `ppt`, `pptx`, `xls`, `xlsx`,
`csv`, `html`, `htm`, `odt`, `rtf`, `pages`, `capl`, `can`, `png`, `jpg`,
`jpeg`, `gif`, `bmp`, `webp`

### Professional Document Preview
- Token-based preview links backed by `preview_data.pkl`.
- Standalone preview route through `preview_token` query parameters.
- Multi-panel preview experience:
  - Viewer
  - Summary
  - Search
  - Q&A
  - Tables
  - Images
  - Downloads
- PDF page rendering in small batches for large manuals.
- Optional PDF table detection on previewed pages.
- PDF zoom control in the preview sidebar.
- Long text pagination with optional full-document mode.
- Spreadsheet sheet selection and row pagination.
- PPTX slide-level text/table inspection.
- Image preview with optional OCR when `pytesseract` is available.
- Downloads for summaries, extracted text, tables, images, and generated DOCX reports.

### Text and Asset Extraction
- PDF text and metadata extraction with `pdfplumber`.
- DOCX paragraphs, headings, tables, metadata, and embedded images.
- PPTX slide text, tables, and embedded images.
- XLSX/CSV structured table extraction.
- HTML text and metadata extraction with BeautifulSoup.
- ODT, RTF, TXT, Markdown, log, CAPL, and CAN text extraction.
- Best-effort recovery for legacy Office files (`doc`, `ppt`, `xls`) and Apple Pages files.
- PNG rendering helpers for extracted tables and images.

### Chat Tab
- Chat with selected workspace documents.
- Retrieval-augmented answers using FAISS, HuggingFace embeddings, and LangChain.
- Local HuggingFace pipeline support through `google/flan-t5-small`.
- Graceful fallback to closest workspace memory when the LLM is unavailable.
- Chat history persisted into workspace memory.
- Direct command support for common document tasks:
  - `summarize` / `analyze`
  - `overview`
  - `find`, `search`, `locate`, `highlight`
  - `count`, `how many`, `number of`, `occurrences`
  - `compare`
  - `pin diagram`, `pin table`, `pin configuration`, `connector details`
  - `item details`, `item information`, `extract item`, `about item`
  - bare item lookups such as `VN1671` or `VN 1671`
- Download support for extracted chat assets such as images, tables, CSV data, and diagrams.

### Dashboard Tab
- Focused dashboard for selected HTML/HTM and XLSX files.
- Extracts useful spreadsheet and report data for visual analysis.
- Interactive Plotly chart support with pie and bar chart modes.
- Login/test/report style data summaries where structured values are available.
- Dashboard-specific file selection independent of other tabs.
- Workspace memory snapshot and AI insight panels are intentionally not shown in the Dashboard tab.

### Compare Tab
- Compare two or more selected files.
- Inline word-level diff rendering.
- Side-by-side line comparison.
- Word presence and difference summaries.
- Semantic comparison summary when available.
- Excel export for comparison results with highlighted differences.
- Works with extracted text from PDFs, Office files, spreadsheets, HTML, text, CAN, and CAPL files.

### CAPL Tab
- CAPL Compiler and Analyzer workspace.
- Select `.can`, `.capl`, or text-based CAPL-like files.
- Create a new CAPL script inside the app.
- Live code preview and static analysis.
- Checks for common CAPL structure, event, variable, syntax, and statement issues.
- Highlighted code and issue summaries.
- AI-assisted fix suggestions and suggested corrected code.
- Autonomous CAPL agent goal input, run history, and workspace memory integration.

### Workspace Memory
- Persistent SQLite workspace storage in `workspace_memory.db`.
- Stores:
  - Indexed files
  - Chat entries
  - Memory events
  - CAPL agent runs
  - Metadata and summaries
- Unified memory text can be embedded into a shared FAISS vector store.
- Memory logs are written to the `workspace_logs` table.

### UI and Responsiveness
- Streamlit branding is hidden by custom CSS.
- Compact header with IntelliDoc AI title, logout, helper button, and status strip.
- Bottom-right helper popup behavior for tab-specific help.
- Sticky random neon glow identity for active main tabs:
  - Chat
  - Dashboard
  - Compare
  - CAPL
- The active tab color is assigned once and stored in `st.session_state.tab_colors`.
- Smooth CSS-only hover, fade, and active-tab transitions.
- Responsive layout for desktop and mobile.

---

## Performance Features

- `CacheManager` LRU cache with TTL support.
- File text cache: up to 100 items.
- Vector store cache: up to 20 items.
- Excel data cache: up to 50 items.
- Embeddings cache: up to 200 items.
- Hash-based change detection prevents unnecessary reprocessing.
- Large PDFs are rendered and scanned lazily.
- Preview persistence uses atomic temp-file replacement when possible.

Important limits:

- `MAX_VECTOR_TEXT_CHARS`: 250,000 characters.
- `PDF_PREVIEW_WINDOW`: 25.
- `PDF_ASSET_SCAN_PAGE_LIMIT`: 10.
- Preview tokens are cleaned up after about 1 hour.

---

## Installation

### Prerequisites
- Python 3.10 or newer.
- A virtual environment is recommended.

### Setup

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

On macOS/Linux:

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### Run

```bash
streamlit run app.py
```

Then open the Streamlit URL shown in the terminal.

---

## Main Dependencies

- Streamlit for the app shell.
- pandas, openpyxl, xlrd, and Plotly for data handling and visualization.
- pdfplumber for PDF text/page extraction.
- python-docx and python-pptx for Office document processing.
- BeautifulSoup for HTML/XML-style extraction.
- Pillow for image and table rendering.
- LangChain, FAISS, sentence-transformers, transformers, torch, and torchvision for local RAG/LLM features.
- pytesseract for optional OCR support.
- SQLite through the Python standard library for workspace memory.

---

## Basic Workflow

1. Log in to the app.
2. Upload files from the sidebar.
3. Select one or more uploaded files.
4. Open a preview when you need full document inspection.
5. Use Chat for natural-language Q&A and direct extraction commands.
6. Use Dashboard for structured HTML/XLSX analysis and charts.
7. Use Compare for diffs across selected files.
8. Use CAPL for CAPL script analysis, suggested fixes, and autonomous runs.
9. Log out to clear the active user workspace state.

---

## Generated Local Data

These files may be created while the app runs:

- `preview_data.pkl`: preview tokens and preview file data.
- `workspace_memory.db`: SQLite workspace memory and logs.
- `active_users.json`: active-user tracking.

These files are runtime state. Keep or delete them depending on whether you want
workspace continuity between app sessions.

---

## Troubleshooting Notes

- If the LLM cannot be loaded, Chat still falls back to retrieved workspace memory.
- If `transformers`, `torch`, or `torchvision` are missing or incompatible, AI generation may be unavailable.
- Large PDFs are intentionally rendered in batches to keep the app responsive.
- Legacy Office and Apple Pages extraction is best-effort; export to DOCX, PPTX, XLSX, PDF, TXT, RTF, ODT, or HTML for better results.
- OCR requires a working Tesseract installation in addition to the `pytesseract` Python package.
- CAPL analysis is static and best-effort; validate generated fixes in CANoe or your official CAPL toolchain.

---

## Contact

For support or questions: vigneshs075@gmail.com
