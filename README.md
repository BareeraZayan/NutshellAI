# 🥜 Nutshell AI — Document Summarizer Agent

**Nutshell AI** is an intelligent, multi-format AI document summarizer and conversational agent built with **Streamlit**, **Groq Cloud LLM**, and **MongoDB Atlas**. 

Upload complex documents (PDF, DOCX, PPTX, TXT), get instant high-fidelity multilingual summaries, query your documents interactively via chat, track token consumption, and export branded PDF reports.

---

## ✨ Features

- **Multi-Format Ingestion**: Supports `.pdf`, `.docx`, `.pptx`, and `.txt` files up to 10MB.
- **Smart Chunking & Hierarchical Summarization**: Handles lengthy documents by splitting them into semantic chunks and generating aggregated synthesis.
- **Multilingual Support**: Generate summaries in English, Urdu, Arabic, and French.
- **Flexible Summary Length**: Choose between **Short** (3 bullet points), **Medium** (5 bullet points), or **Detailed** (10 comprehensive points).
- **Interactive Conversational Agent**: Ask follow-up questions about the document with full context awareness.
- **Persistent Session Storage**: Save, browse, and resume previous document analysis sessions via MongoDB Atlas.
- **PDF Report Export**: Download cleanly formatted PDF summary reports directly from the app.
- **Token Budget & Usage Tracking**: Live daily token budget counter and real-time usage metrics per request.

---

## 🏗️ Architecture Overview

```
 ┌─────────────────┐       ┌────────────────────────┐
 │   User Upload   │ ────> │  Format Extraction     │
 │ (PDF/DOCX/PPTX) │       │ (PyPDF2, docx, pptx)   │
 └─────────────────┘       └───────────┬────────────┘
                                       │ Raw Text
                                       ▼
 ┌─────────────────┐       ┌────────────────────────┐
 │  MongoDB Atlas  │ <───> │  Groq LLM Engine       │
 │ (Session State) │       │ (Chunking & Synthesis) │
 └─────────────────┘       └───────────┬────────────┘
                                       │
                                       ▼
                       ┌───────────────────────────────┐
                       │    Streamlit Modern Web UI    │
                       │ - Multilingual Summary View   │
                       │ - Interactive Document Chat   │
                       │ - PDF Export & Token Tracker  │
                       └───────────────────────────────┘
```

---

## 🚀 Quick Start

### 1. Clone the Repository
```bash
git clone https://github.com/BareeraZayan/NutshellAI.git
cd NutshellAI
```

### 2. Create and Activate a Virtual Environment
```bash
# Windows
python -m venv venv
.\venv\Scripts\activate

# macOS / Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure Environment Variables
Copy `.env.example` to `.env` and fill in your credentials:
```bash
cp .env.example .env
```

Edit `.env`:
```env
GROQ_API_KEY=gsk_your_groq_api_key_here
GROQ_MODEL=openai/gpt-oss-120b
MONGO_URI=mongodb+srv://<username>:<password>@cluster0.example.mongodb.net/?retryWrites=true&w=majority
```

### 5. Verify Database Connection
```bash
python test_connection.py
```

### 6. Run the Application
```bash
streamlit run app.py
```

Open your browser at `http://localhost:8501`.

---

## 📁 Project Structure

```text
├── app.py              # Main Streamlit application and core pipeline
├── test_connection.py  # MongoDB Atlas diagnostic utility
├── requirements.txt    # Python package dependencies
├── .env.example        # Environment variable template
├── .gitignore          # Git exclusion rules for secrets and caches
└── README.md           # Project documentation
```

---

## ⚙️ Configuration & Tech Stack

| Component | Technology | Purpose |
|---|---|---|
| **Frontend** | Streamlit | Responsive web UI, session state, and metric widgets |
| **LLM Inference** | Groq Cloud | Ultra-fast inference (`openai/gpt-oss-120b`) |
| **Database** | MongoDB Atlas / PyMongo | Chat session history and document metadata storage |
| **File Extraction** | PyPDF2, python-docx, python-pptx | Parsing multi-format enterprise files |
| **Document Export** | FPDF2 | Automated summary PDF report generation |

---

## 🛡️ License

This project is licensed under the MIT License.
