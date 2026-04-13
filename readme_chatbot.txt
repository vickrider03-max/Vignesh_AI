# 🧠 IntelliDoc AI – Smart Document Assistant

---

## 🚀 Overview

**IntelliDoc AI** is a powerful **multi-utility document analyzer** built with Streamlit.  
It enables users to **upload, analyze, compare, and interact with files using AI**, along with specialized support for **CAPL script analysis and auto-fixing**.

---

## 🧩 Features

### 📂 File Management
- Upload multiple formats:
  `PDF, DOCX, PPTX, XLSX, TXT, HTML, CAPL (.can)`
- Multi-file selection & filtering
- Persistent preview system

---

### 🔍 Smart Document Preview
- Open preview in new tab
- Keyword highlighting
- Page-level PDF preview
- Extract:
  - Text
  - Tables
  - Images
  - Metadata

---

### 🧠 AI Chat (RAG System)
- Ask questions about uploaded files
- Context-aware responses
- Multi-file semantic understanding
- Powered by:
  - FAISS (vector DB)
  - HuggingFace models

---

### 📊 Dashboard & Analytics
- Excel/CSV visualization
- Trends & statistics
- Interactive charts (Plotly)
- Export insights

---

### 🔄 File Comparison
- Compare 2+ files
- Word-level diff
- Inline visual comparison
- Export results to Excel

---

### 🚗 CAPL Script Analyzer
- Upload or create `.can` files
- Built-in CAPL editor
- Code analysis & issue detection
- Suggestions & improvements

---

### 🤖 AI CAPL Auto-Fix
```text
Analyze → Suggest Fix → Apply Fix → Save

🧭 Application Flow
Sidebar:
+----------------------+
| Upload Files         |
| Select Files         |
| Filter CAPL (.can)   |
+----------------------+

Tabs:
[ Chat ] → AI Q&A  
[ Dashboard ] → Analytics  
[ Compare ] → Diff  
[ CAPL ] → Editor + AI Fix  

🛠️ CAPL Workflow
Create/Edit Script
        ↓
Analyze Code
        ↓
AI Suggest Fix?
    ↓        ↓
   Yes       No
    ↓         ↓
Apply Fix    Save
    ↓
Update Editor


⚠️ Notes
CAPL files must use .can extension
AI features require configured models/API
Limit large file comparisons for better performance

🙌 Credits
Built with ❤️ using Streamlit
AI powered by HuggingFace + LangChain
Visualization via Plotly


📧 Contact
📩 vigneshs075@gmail.com
