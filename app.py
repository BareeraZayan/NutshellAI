import streamlit as st
from groq import Groq
from PyPDF2 import PdfReader
from docx import Document
from pptx import Presentation
from pymongo import MongoClient
from dotenv import load_dotenv
from fpdf import FPDF
import os
import uuid

load_dotenv()

client = Groq(api_key=os.getenv("GROQ_API_KEY"))
GROQ_MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-120b")
mongo_client = MongoClient(os.getenv("MONGO_URI"))
db = mongo_client["summarizer_db"]
collection = db["chat_sessions"]

DAILY_TOKEN_LIMIT = 100000
MAX_FILE_SIZE_MB = 10

LANG_MAP = {"English": "in English", "Urdu": "in Urdu", "Arabic": "in Arabic", "French": "in French"}

# ---------------- Extraction Functions ----------------

def extract_text_from_pdf(file):
    reader = PdfReader(file)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    return text

def extract_text_from_docx(file):
    doc = Document(file)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def extract_text_from_pptx(file):
    prs = Presentation(file)
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        text += run.text + " "
    return text

def extract_text(uploaded_file):
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(uploaded_file)
    elif ext == ".docx":
        return extract_text_from_docx(uploaded_file)
    elif ext == ".pptx":
        return extract_text_from_pptx(uploaded_file)
    elif ext == ".txt":
        return uploaded_file.read().decode("utf-8")
    else:
        return None

# ---------------- Summarization ----------------

def summarize_document(text, language="English", length="Medium"):
    max_chars = 15000
    lang_instruction = LANG_MAP.get(language, "in English")

    if length == "Short":
        bullet_instruction = "3 bullet points"
    elif length == "Detailed":
        bullet_instruction = "10 detailed bullet points"
    else:
        bullet_instruction = "5 bullet points"

    if len(text) <= max_chars:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": f"Summarize this document in {bullet_instruction}, {lang_instruction}:\n\n{text}"}]
        )
        return response.choices[0].message.content, response.usage.total_tokens
    else:
        chunks = [text[i:i+max_chars] for i in range(0, len(text), max_chars)]
        chunk_summaries = []
        total_tokens = 0
        for chunk in chunks:
            response = client.chat.completions.create(
                model=GROQ_MODEL,
                messages=[{"role": "user", "content": f"Summarize this section in 3-4 bullet points, {lang_instruction}:\n\n{chunk}"}]
            )
            chunk_summaries.append(response.choices[0].message.content)
            total_tokens += response.usage.total_tokens
        combined = "\n\n".join(chunk_summaries)
        final_response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": f"Combine these section summaries into {bullet_instruction}, {lang_instruction}:\n\n{combined}"}]
        )
        total_tokens += final_response.usage.total_tokens
        return final_response.choices[0].message.content, total_tokens

# ---------------- Chat (Auto-detects language, blocks Hindi) ----------------

def chat_with_document(document_text, user_question, chat_history):
    base_instruction = (
        "Your name is Nutshell AI. If the user asks your name, tell them your name is Nutshell AI. "
        "If the user asks who made you, who created you, who is your founder, or who is your developer, "
        "tell them you were created by Bareera Zayan. Do not mention a 'team of developers' or any other name. "
        "PERSONALITY: Be warm, clear, and professional. Do not use emojis in your replies. "
        "You are not limited to document questions — you can chat with the user about anything, just like a "
        "knowledgeable, professional assistant would. Only focus on the uploaded document when the user's "
        "question is actually about it. "
        "IMPORTANT LANGUAGE RULES: Detect the language the user is typing in and reply in that same language. "
        "Supported reply languages are: English, Urdu (Urdu script), Roman Urdu (Urdu words in English letters), Arabic, and French. "
        "If the user writes in Hindi or Devanagari script, do NOT reply in Hindi — instead reply in Roman Urdu or English. "
        "Never produce any Hindi or Devanagari script text under any circumstance, even if asked directly."
    )
    if document_text:
        system_content = base_instruction + " You are a helpful assistant answering questions about the following document(s). Use this as your main source of information.\n\nDocument:\n" + document_text
    else:
        system_content = base_instruction + " No document has been uploaded yet, so answer as a general helpful assistant."

    context_messages = [{"role": "system", "content": system_content}]
    context_messages.extend(chat_history)
    context_messages.append({"role": "user", "content": user_question})
    response = client.chat.completions.create(model=GROQ_MODEL, messages=context_messages)
    return response.choices[0].message.content, response.usage.total_tokens

# ---------------- Document Comparison ----------------

def compare_documents(text1, text2, language="English"):
    lang_instruction = LANG_MAP.get(language, "in English")
    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": f"Compare these two documents and list key similarities and differences in bullet points, {lang_instruction}:\n\nDocument 1:\n{text1[:7000]}\n\nDocument 2:\n{text2[:7000]}"}]
    )
    return response.choices[0].message.content, response.usage.total_tokens

# ---------------- PDF Export (Unicode-safe) ----------------

# Put a Unicode TTF font next to app.py for this to work with Urdu/Arabic/French text.
# Recommended: download "DejaVuSans.ttf" (Latin/Cyrillic/Greek, good for English/French)
# and, if you need Urdu/Arabic glyphs rendered too, also grab "NotoNastaliqUrdu-Regular.ttf"
# or "NotoNaskhArabic-Regular.ttf" from Google Fonts and place them in the same folder.
FONT_PATH = os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf")

def create_pdf(text_content):
    pdf = FPDF()
    pdf.add_page()

    if os.path.exists(FONT_PATH):
        pdf.add_font("DejaVu", "", FONT_PATH)
        pdf.set_font("DejaVu", size=12)
        pdf.multi_cell(0, 10, text_content)
    else:
        # Fallback: strip characters the built-in font can't render instead of crashing
        pdf.set_font("Arial", size=12)
        safe_text = text_content.encode("latin-1", errors="ignore").decode("latin-1")
        pdf.multi_cell(0, 10, safe_text)

    output = pdf.output()
    if isinstance(output, str):
        # Classic fpdf (not fpdf2) returns a plain string here — encode it ourselves
        return output.encode("latin-1")
    return bytes(output)

def format_chat_for_export(chat_history):
    lines = []
    for msg in chat_history:
        role = "You" if msg["role"] == "user" else "Nutshell AI"
        lines.append(f"{role}: {msg['content']}\n")
    return "\n".join(lines)

# ---------------- MongoDB ----------------

def load_history(session_id):
    record = collection.find_one({"_id": session_id})
    if record:
        return record.get("chat_history", []), record.get("document_text", None)
    return [], None

def save_history(session_id, chat_history, document_text):
    collection.update_one({"_id": session_id}, {"$set": {"chat_history": chat_history, "document_text": document_text}}, upsert=True)

def log_analytics(event_type):
    collection.update_one({"_id": "analytics"}, {"$inc": {f"counts.{event_type}": 1}}, upsert=True)

def get_analytics():
    record = collection.find_one({"_id": "analytics"})
    if record:
        return record.get("counts", {})
    return {}

# ---------------- Streamlit UI ----------------

st.set_page_config(
    page_title="Document Summarizer Assistant",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

theme_css = """
<style>
:root {
    --void: #0a0f1c;
    --panel: rgba(16, 26, 46, 0.88);
    --panel-strong: rgba(22, 33, 58, 0.96);
    --border: rgba(238, 241, 255, 0.12);
    --text: #eef1ff;
    --muted: #93a0c4;
    --cobalt: #4f74ff;
    --cobalt-bright: #7fa0ff;
    --cyan: #2dd4bf;
    --amber-warn: #f0a840;
    --danger: #ff6b6b;
}
html, body, .stApp {
    background: radial-gradient(circle at top left, #16213a 0%, #0a0f1c 45%, #05070d 100%) !important;
    color: var(--text) !important;
}
[data-testid="stAppViewContainer"] {
    background: transparent !important;
}
.stApp * {
    color: var(--text) !important;
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, rgba(8, 13, 24, 0.98), rgba(14, 22, 40, 0.98)) !important;
    border-right: 1px solid var(--border) !important;
}
.stMetric {
    background: rgba(255,255,255,0.04);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 0.7rem 0.8rem;
    margin-bottom: 0.5rem;
}

/* ---- File uploader: force readable text on the dropzone ---- */
section[data-testid="stFileUploader"] {
    border: 1px solid var(--border) !important;
    border-radius: 18px !important;
    background: linear-gradient(135deg, rgba(79,116,255,0.12), rgba(45,212,191,0.08)) !important;
    padding: 0.35rem !important;
}
section[data-testid="stFileUploader"] section {
    background: rgba(10, 15, 28, 0.55) !important;
    border-radius: 14px !important;
}
section[data-testid="stFileUploader"] * {
    color: var(--text) !important;
}
section[data-testid="stFileUploader"] small {
    color: var(--muted) !important;
}
section[data-testid="stFileUploader"] button {
    border-radius: 999px !important;
    background: linear-gradient(90deg, var(--cobalt-bright), var(--cobalt)) !important;
    color: #050914 !important;
    border: none !important;
    font-weight: 700 !important;
}
section[data-testid="stFileUploader"] button * {
    color: #050914 !important;
}
/* Hide Streamlit's native uploaded-file chip — we render our own styled card instead */
.stApp [data-testid="stFileUploaderFile"],
.stApp [data-testid^="stFileUploaderFile"],
.stApp [data-testid="stFileUploaderFiles"],
.stApp [class*="uploadedFile" i],
.stApp section[data-testid="stFileUploader"] ul,
.stApp section[data-testid="stFileUploader"] li,
.stApp section[data-testid="stFileUploader"] > div > div:nth-child(2) {
    display: none !important;
}
/* keep the dropzone (drag & drop area + Browse button) itself visible */
.stApp section[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzone"] {
    display: flex !important;
}

button, .stButton > button {
    border-radius: 999px !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
    background: linear-gradient(90deg, var(--cobalt-bright), var(--cobalt)) !important;
    color: #050914 !important;
    font-weight: 700 !important;
    box-shadow: 0 10px 25px rgba(79,116,255,0.20) !important;
    min-height: 44px !important;
    padding: 0.55rem 1.4rem !important;
    white-space: nowrap !important;
}
button:hover, .stButton > button:hover {
    filter: brightness(1.08) !important;
}
.stButton {
    min-width: fit-content !important;
}
.stApp [data-testid="stChatInput"],
.stApp [data-testid="stChatInput"] > div,
.stApp [data-testid="stChatInput"] textarea,
.stApp [data-testid="stBottomBlockContainer"] [data-testid="stChatInput"] {
    background-color: rgba(16, 26, 46, 0.95) !important;
    border: 1px solid var(--border) !important;
    border-radius: 18px !important;
    color: var(--text) !important;
    -webkit-text-fill-color: var(--text) !important;
    caret-color: var(--text) !important;
}
.stApp [data-testid="stChatInput"] textarea::placeholder {
    color: var(--muted) !important;
    -webkit-text-fill-color: var(--muted) !important;
    opacity: 1 !important;
}
.stApp [data-testid="stExpander"] {
    background: var(--panel) !important;
    border: 1px solid var(--border) !important;
    border-radius: 16px !important;
}
.stApp [data-testid="stExpander"] summary {
    background: transparent !important;
    color: var(--text) !important;
    font-weight: 700 !important;
}
.stApp [data-testid="stExpander"] details {
    background: transparent !important;
}
[data-testid="stChatMessage"] {
    background-color: rgba(16, 26, 46, 0.85) !important;
    border: 1px solid var(--border) !important;
    border-radius: 16px !important;
    padding: 0.75rem !important;
    margin-bottom: 0.6rem !important;
}
[data-testid="stChatMessage"] p,
[data-testid="stChatMessage"] span,
[data-testid="stChatMessage"] div {
    color: var(--text) !important;
}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) {
    background-color: rgba(79, 116, 255, 0.14) !important;
    border-color: rgba(79, 116, 255, 0.3) !important;
}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"]) {
    background-color: rgba(45, 212, 191, 0.10) !important;
    border-color: rgba(45, 212, 191, 0.25) !important;
}
.stTabs [data-testid="stBaseButtonHeader"] {
    background: rgba(255,255,255,0.04) !important;
    border: 1px solid var(--border) !important;
    border-radius: 999px !important;
}
.stTabs [data-testid="stBaseButtonHeader"] button {
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
}
.page-title {
    font-size: 1.9rem;
    font-weight: 800;
    margin: 0.3rem 0 0.1rem 0;
    color: var(--text) !important;
}
.page-subtitle {
    color: var(--muted) !important;
    font-size: 0.95rem;
    margin: 0 0 1.1rem 0;
}
.hero-card {
    margin: 0.3rem 0 1rem 0;
    padding: 1.3rem 1.4rem;
    border-radius: 24px;
    background: linear-gradient(135deg, rgba(79,116,255,0.16), rgba(45,212,191,0.10));
    border: 1px solid var(--border);
    box-shadow: 0 20px 45px rgba(0,0,0,0.30);
}
.hero-card h1 {
    font-size: 2rem;
    margin: 0.2rem 0 0.4rem 0;
    font-weight: 800;
}
.hero-card p {
    color: var(--muted) !important;
    margin-bottom: 0.4rem;
}
.hero-badge {
    display: inline-block;
    padding: 0.35rem 0.7rem;
    border-radius: 999px;
    background: rgba(255,255,255,0.08);
    color: #c9d6ff !important;
    font-size: 0.82rem;
    font-weight: 700;
    letter-spacing: 0.04em;
    text-transform: uppercase;
}
.hero-pill {
    display: inline-block;
    margin: 0.45rem 0.4rem 0 0;
    padding: 0.4rem 0.7rem;
    border-radius: 999px;
    background: rgba(255,255,255,0.06);
    color: #dfe6ff !important;
    border: 1px solid rgba(255,255,255,0.08);
}
.info-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 0.8rem;
    margin-bottom: 1rem;
}
.info-card {
    padding: 0.95rem;
    border-radius: 18px;
    background: var(--panel);
    border: 1px solid var(--border);
    backdrop-filter: blur(14px);
}
.info-card h4 {
    margin: 0.2rem 0 0.25rem 0;
    font-size: 1rem;
}
.info-card p {
    color: var(--muted) !important;
    font-size: 0.92rem;
    margin: 0;
}
.info-icon {
    font-size: 1.2rem;
    margin-bottom: 0.3rem;
}
.section-heading {
    font-size: 1.02rem;
    font-weight: 700;
    margin-bottom: 0.6rem;
    color: #dbe4ff !important;
}
.control-panel {
    padding: 0.8rem;
    border-radius: 18px;
    background: var(--panel);
    border: 1px solid var(--border);
    margin-bottom: 0.8rem;
}
.sidebar-brand {
    display: flex;
    align-items: center;
    gap: 0.7rem;
    margin-bottom: 0.8rem;
