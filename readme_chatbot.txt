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
- Uploaded files persist in the active session and are remembered per tab.
- Files are not auto-selected after upload; users explicitly choose which files each tab should use.
- Selected files can still be shared across Chat, Dashboard, Compare, and CAPL tabs when selected by the user.
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
- Enterprise Document Intelligence Agent for selected workspace documents.
- The chat system behaves like an intent-routing agent rather than a single generic prompt.
- Supports grounded chat over PDF, DOCX, PPTX, XLSX, TXT, RTF, ODT, HTML, and HTM files.
- Classifies each request into one primary intent:
  - `FULL_DOCUMENT_ANALYSIS`
  - `SHORT_SUMMARY`
  - `OVERVIEW`
  - `QUESTION_ANSWERING`
  - `TABLE_EXTRACTION`
  - `COMPONENT_EXTRACTION`
  - `DIAGRAM_ANALYSIS`
  - `COMPARISON`
  - `SEARCH`
- Three clearly separated chat pipelines:
  - Document Analysis: `Analyze` uses map-reduce summarization, not vector search.
  - Question Answering: specific questions use hybrid RAG/vector retrieval.
  - Fast Summary Mode: `Summary` uses cached summary/analysis for instant responses where available.
- Map-reduce analysis flow:
  - Splits meaningful document text into 1,000-word chunks with 150-word overlap.
  - Processes chunk summaries with the production chunk-analysis prompt.
  - Runs a final reduce prompt to produce the full structured document analysis.
  - Uses `ThreadPoolExecutor(max_workers=6)` for parallel chunk processing when the local LLM is available.
- Chat summary and analysis cache:
  - Stored in `st.session_state.doc_cache`.
  - Mirrored through `st.session_state.document_summary` for compatibility.
  - Cleared on single-file delete, Clear All, and logout.
- Hybrid retrieval-augmented question answering uses FAISS dense search, BM25 keyword search, HuggingFace embeddings, LangChain, and reranking.
- Answers are grounded only in uploaded document content.
- The prompt tells the assistant to read all retrieved context, combine partial evidence, and avoid refusing too early.
- RAG question answers should say `Not specified in the provided context` when the answer is missing from retrieved context.
- Vague document commands are routed directly:
  - `Analyze` -> full document analysis.
  - `Summary` -> short summary.
  - `Overview` -> high-level overview.
- Dedicated intent routes are available for:
  - table/data extraction
  - component/module extraction
  - diagram/flow analysis
  - comparison
  - search/find/count
- These direct commands ignore cover pages, metadata, table-of-contents lines, repeated headers/footers, and dotted TOC entries as much as possible.
- Query-time retrieval flow:
  - Optional LLM query rewrite to improve document search while preserving the original intent.
  - Dense semantic retrieval from FAISS.
  - Sparse keyword retrieval from BM25.
  - Merge and deduplicate dense + sparse candidates.
  - Rerank with `CrossEncoder` when available.
  - Fall back to lexical reranking if a cross-encoder model cannot be loaded.
  - Send the top 5-8 reranked chunks to the LLM with citation metadata.
- Every supported file is normalized into structured records:
  - `text`
  - `metadata.file_name`
  - `metadata.file_type`
  - `metadata.page_or_sheet`
  - `metadata.section`
  - `metadata.document_id`
  - `metadata.chunk_index`
- Metadata-preserving retrieval keeps each chunk traceable to its original PDF page, slide, sheet, section, or file.
- Source citations are included in answers using format-aware labels:
  - `file_name (PDF Page 3)`
  - `file_name (Sheet: Sales_Q1)`
  - `file_name (Slide 5)`
  - `file_name (Section: Introduction)`
  - `file_name` when no locator is available.
- Excel sheets are treated as structured data sources for row, column, and value reasoning when relevant rows are retrieved.
- PowerPoint slides are treated as sections.
- HTML is cleaned with BeautifulSoup before retrieval so scripts, styles, navigation, and footer noise are removed.
- Per-user, per-document-selection memory keeps follow-up questions scoped to the same selected documents.
- Visible chat transcripts are also scoped by selected document set.
- Every chat response includes a confidence footer:
  - `Confidence: High`
  - `Confidence: Medium`
  - `Confidence: Low`
  Confidence is inferred from available context, citations, and missing-information signals.
- FAISS ChatPDF vector stores are persisted locally under `chatpdf_vectorstores/`.
- Local HuggingFace pipeline support through `google/flan-t5-small`.
- Graceful extractive fallback from retrieved chunks when the LLM is unavailable.
- Streaming-style assistant rendering gives a ChatGPT-like typing effect in the Streamlit UI.
- Direct command support for deterministic document checks:
  - `find`, `search`, `locate`, `highlight`
  - `count`, `how many`, `number of`, `occurrences`
- Download support for extracted chat assets such as images, tables, CSV data, and diagrams.

### Dashboard Tab
- Focused dashboard for selected HTML/HTM and XLSX files.
- Extracts useful spreadsheet and report data for visual analysis.
- Interactive Plotly chart support with pie and bar chart modes.
- Login/test/report style data summaries where structured values are available.
- Dashboard-specific file selection independent of other tabs.
- Dashboard helper guidance now focuses on HTML/HTM/XLSX compatibility, selected source context, chart type, orientation, and reset behavior.
- Workspace memory snapshot and AI insight panels are intentionally not shown in the Dashboard tab.

### Compare Tab
- Compare two or more selected files.
- Inline word-level diff rendering.
- Side-by-side line comparison.
- Word presence and difference summaries.
- Semantic comparison summary when available.
- Excel export for comparison results with highlighted differences.
- Works with extracted text from PDFs, Office files, spreadsheets, HTML, text, CAN, and CAPL files.
- Compare helper guidance explains when to use the Compare tab for exact file diffs and when to use Chat comparison intent for natural-language comparison answers.

### CAPL Tab
- CAPL Compiler and Analyzer workspace.
- Select `.can`, `.capl`, or text-based CAPL-like files.
- Create a new CAPL script inside the app.
- Live code preview and static analysis.
- Checks for common CAPL structure, event, variable, syntax, and statement issues.
- Highlighted code and issue summaries.
- AI-assisted fix suggestions and suggested corrected code.
- Autonomous CAPL agent goal input, run history, and workspace memory integration.
- CAPL helper guidance emphasizes selecting `.can` or `.txt` sources, reviewing static issues, using AI fix suggestions, and validating fixes in official CAPL tooling.

### Workspace Memory
- Persistent SQLite workspace storage in `workspace_memory.db`.
- Stores:
  - Indexed files
  - Chat entries
  - Memory events
  - CAPL agent runs
  - Metadata and summaries
- Unified memory text can be embedded into a shared FAISS vector store.
- ChatPDF conversation memory is held per user and selected document set during the active session.
- ChatPDF FAISS collections persist on disk by user/document-selection signature.
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
- ChatPDF collections are keyed by user id plus selected file names and file hashes.
- ChatPDF collections include a schema version in the key so citation metadata changes rebuild stale indexes safely.
- ChatPDF chunks preserve source metadata so every retrieved chunk can be traced back to its document page, slide, sheet, section, or file-level record.
- BM25 sparse indexes are cached in session state per user/document-selection.
- Structured document chunks are cached in session state per user/document-selection for reuse by FAISS, BM25, and reranking.
- Full-document analysis and fast summaries are cached in `st.session_state.doc_cache` so `Analyze` and `Summary` do not repeatedly process the same document text.
- Map-reduce summarization uses standard-library `concurrent.futures.ThreadPoolExecutor` for parallel chunk processing when an LLM is available.
- Optional cross-encoder reranking uses `CHATPDF_RERANKER_MODEL` when set; otherwise it defaults to `cross-encoder/ms-marco-MiniLM-L-6-v2`.
- Optional query rewrite reuses the local LLM when available; retrieval still runs with the original user question if rewriting fails or the LLM is unavailable.
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
- LangChain, FAISS, rank-bm25, sentence-transformers, transformers, torch, and torchvision for hybrid local RAG/LLM features.
- pytesseract for optional OCR support.
- SQLite through the Python standard library for workspace memory.

Note: the current implementation is a Streamlit + FAISS application. It does not require FastAPI, ChromaDB, OpenAI SDK, Redis, S3 SDKs, or JWT packages unless a separate service architecture is added later.

---

## Basic Workflow

1. Log in to the app.
2. Upload files from the sidebar.
3. Explicitly select one or more uploaded files for the tab you want to use.
4. Open a preview when you need full document inspection.
5. Use Chat for document-only Q&A with source citations.
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
- `chatpdf_vectorstores/`: persisted FAISS indexes for page-aware ChatPDF retrieval.

These files are runtime state. Keep or delete them depending on whether you want
workspace continuity between app sessions.

---

## Troubleshooting Notes

- If the LLM cannot be loaded, Chat falls back to extractive answers from retrieved document chunks.
- If Chat says information is unavailable, the final reranked context was empty or unrelated to the question.
- If retrieval quality feels weak, install dependencies from `requirements.txt` and allow the optional cross-encoder reranker model to load, or set `CHATPDF_RERANKER_MODEL` to a locally available reranker.
- Citations are only as precise as the source format allows: PDFs use page numbers, PPTX uses slide numbers, XLSX uses sheet names, and long text/Word/HTML-like files use generated section/page labels.
- If `transformers`, `torch`, or `torchvision` are missing or incompatible, AI generation may be unavailable.
- Large PDFs are intentionally rendered in batches to keep the app responsive.
- Legacy Office and Apple Pages extraction is best-effort; export to DOCX, PPTX, XLSX, PDF, TXT, RTF, ODT, or HTML for better results.
- OCR requires a working Tesseract installation in addition to the `pytesseract` Python package.
- CAPL analysis is static and best-effort; validate generated fixes in CANoe or your official CAPL toolchain.

---

## Contact

For support or questions: vigneshs075@gmail.com
