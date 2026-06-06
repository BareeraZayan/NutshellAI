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
}
.sidebar-brand .brand-icon {
    display: grid;
    place-items: center;
    width: 42px;
    height: 42px;
    border-radius: 14px;
    background: linear-gradient(135deg, var(--cobalt-bright), var(--cobalt));
    font-size: 1.15rem;
}
.sidebar-brand h3 {
    margin: 0;
    font-size: 1rem;
}
.sidebar-brand p {
    margin: 0;
    color: var(--muted) !important;
    font-size: 0.82rem;
}
.sidebar-card {
    padding: 0.8rem;
    border-radius: 16px;
    background: rgba(255,255,255,0.05);
    border: 1px solid var(--border);
    margin-bottom: 0.8rem;
}
.sidebar-card h4 {
    margin: 0 0 0.2rem 0;
    font-size: 0.95rem;
}
.sidebar-card p {
    margin: 0;
    color: var(--muted) !important;
    font-size: 0.85rem;
}
.file-loaded-card {
    display: flex;
    align-items: center;
    gap: 12px;
    background: rgba(45, 212, 191, 0.08);
    border: 1px solid rgba(45, 212, 191, 0.3);
    border-radius: 14px;
    padding: 0.6rem 1rem;
    margin: 0.3rem 0 0.8rem 0;
}
.file-loaded-card .ficon {
    font-size: 1.3rem;
}
.file-loaded-card b {
    color: var(--text) !important;
    font-size: 0.92rem;
}
.file-loaded-card span {
    display: block;
    color: var(--muted) !important;
    font-size: 0.78rem;
}

/* ---- Dynamic top usage banner ---- */
.usage-banner {
    position: sticky;
    top: 0;
    z-index: 999;
    background: linear-gradient(90deg, rgba(240,168,64,0.18), rgba(240,168,64,0.08));
    border: 1px solid rgba(240,168,64,0.35);
    border-radius: 14px;
    backdrop-filter: blur(10px);
    padding: 0.65rem 1rem;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 10px;
    font-size: 0.85rem;
    color: #ffd9a0 !important;
    margin-bottom: 0.8rem;
}
.usage-banner b {
    color: var(--amber-warn) !important;
    font-weight: 700;
}
.usage-banner.danger {
    background: linear-gradient(90deg, rgba(255,107,107,0.18), rgba(255,107,107,0.08));
    border-color: rgba(255,107,107,0.4);
    color: #ffc2c2 !important;
}
.usage-banner.danger b {
    color: var(--danger) !important;
}

/* ---- Sidebar token analytics ---- */
.side-quota {
    background: rgba(240,168,64,0.08);
    border: 1px solid rgba(240,168,64,0.3);
    border-radius: 14px;
    padding: 14px;
    margin-bottom: 14px;
}
.side-quota .qtop {
    display: flex;
    justify-content: space-between;
    align-items: baseline;
    font-size: 0.75rem;
    color: var(--muted) !important;
    margin-bottom: 8px;
}
.side-quota .qtop b {
    color: var(--amber-warn) !important;
    font-size: 0.95rem;
}
.quota-track {
    height: 6px;
    border-radius: 100px;
    background: rgba(255,255,255,0.08);
    overflow: hidden;
    margin-bottom: 10px;
}
.quota-fill {
    height: 100%;
    background: linear-gradient(90deg, var(--amber-warn), var(--danger));
    border-radius: 100px;
}
.quota-count {
    font-size: 0.72rem;
    color: var(--muted) !important;
    display: flex;
    justify-content: space-between;
    margin-bottom: 12px;
}
.quota-count b {
    color: var(--text) !important;
}
.qb-row {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 0.72rem;
    color: var(--muted) !important;
    margin-bottom: 6px;
}
.qb-label {
    width: 74px;
    flex-shrink: 0;
}
.qb-track {
    flex: 1;
    height: 5px;
    border-radius: 100px;
    background: rgba(255,255,255,0.07);
    overflow: hidden;
}
.qb-fill {
    height: 100%;
    border-radius: 100px;
}
.qb-fill.summarize { background: var(--cobalt-bright); }
.qb-fill.chat { background: var(--cyan); }
.qb-fill.compare { background: var(--amber-warn); }
.qb-val {
    width: 42px;
    text-align: right;
    color: var(--muted) !important;
}
</style>"""

st.markdown(theme_css, unsafe_allow_html=True)

# ---------------- Session State ----------------

if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())
if "chat_history" not in st.session_state:
    loaded_chat, loaded_doc = load_history(st.session_state.session_id)
    st.session_state.chat_history = loaded_chat
    st.session_state.document_text = loaded_doc
if "tokens_used_today" not in st.session_state:
    st.session_state.tokens_used_today = 0
if "tokens_by_category" not in st.session_state:
    st.session_state.tokens_by_category = {"summarize": 0, "chat": 0, "compare": 0}
if "current_summary" not in st.session_state:
    st.session_state.current_summary = None
if "compare_document_text" not in st.session_state:
    st.session_state.compare_document_text = None
if "current_comparison" not in st.session_state:
    st.session_state.current_comparison = None

# ---------------- Dynamic usage banner + sidebar analytics ----------------
# NOTE: these are called directly in the main script flow (not inside any
# @st.fragment), so every action anywhere in the app — summarizing, comparing,
# or chatting — triggers a full rerun and these numbers update immediately.

def render_usage_banner():
    tokens_used = st.session_state.tokens_used_today
    usage_percent = (tokens_used / DAILY_TOKEN_LIMIT) * 100
    if usage_percent >= 70:
        css_class = "danger" if usage_percent >= 90 else ""
        st.markdown(f"""
        <div class="usage-banner {css_class}">
            <span>You have used <b>{tokens_used:,} / {DAILY_TOKEN_LIMIT:,}</b> tokens today ({usage_percent:.0f}%) —
            summaries and chat may pause once the limit is reached.</span>
        </div>
        """, unsafe_allow_html=True)

def render_token_analytics():
    cat = st.session_state.tokens_by_category
    total_used = st.session_state.tokens_used_today
    usage_percent = min((total_used / DAILY_TOKEN_LIMIT) * 100, 100)

    def pct(n):
        return min((n / DAILY_TOKEN_LIMIT) * 100, 100)

    st.markdown(f"""
    <div class="side-quota">
        <div class="qtop"><span>Token usage today</span><b>{usage_percent:.0f}%</b></div>
        <div class="quota-track"><div class="quota-fill" style="width:{usage_percent:.0f}%"></div></div>
        <div class="quota-count"><span>Used</span><b>{total_used:,} / {DAILY_TOKEN_LIMIT:,}</b></div>
        <div class="qb-row"><span class="qb-label">Summarize</span><div class="qb-track"><div class="qb-fill summarize" style="width:{pct(cat['summarize']):.0f}%"></div></div><span class="qb-val">{cat['summarize']:,}</span></div>
        <div class="qb-row"><span class="qb-label">Chat</span><div class="qb-track"><div class="qb-fill chat" style="width:{pct(cat['chat']):.0f}%"></div></div><span class="qb-val">{cat['chat']:,}</span></div>
        <div class="qb-row"><span class="qb-label">Compare</span><div class="qb-track"><div class="qb-fill compare" style="width:{pct(cat['compare']):.0f}%"></div></div><span class="qb-val">{cat['compare']:,}</span></div>
    </div>
    """, unsafe_allow_html=True)

render_usage_banner()

with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <div class="brand-icon">✦</div>
        <div>
            <h3>Nutshell AI</h3>
            <p>Document Intelligence Suite</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    render_token_analytics()

    st.markdown("""
    <div class="sidebar-card">
        <h4>Analytics Dashboard</h4>
        <p>Monitor how your workspace is being used.</p>
    </div>
    """, unsafe_allow_html=True)
    stats = get_analytics()
    st.metric("Documents Summarized", stats.get("summarize", 0))
    st.metric("Chat Messages Sent", stats.get("chat", 0))
    st.metric("Documents Compared", stats.get("compare", 0))

st.markdown("""
<div class="page-title">📄 Document Summarizer Assistant</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-card">
    <div class="hero-badge">Premium AI Workspace</div>
    <h1>Turn documents into strategy-ready intelligence.</h1>
    <p>Summarize large files, chat in multiple languages, and compare documents with a polished experience built for modern workflows.</p>
    <div>
        <span class="hero-pill">📄 Multi-format upload</span>
        <span class="hero-pill">💬 Multilingual chat</span>
        <span class="hero-pill">📊 PDF export</span>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="info-grid">
    <div class="info-card">
        <div class="info-icon">⚡</div>
        <h4>Instant summaries</h4>
        <p>Generate sharp summaries in seconds.</p>
    </div>
    <div class="info-card">
        <div class="info-icon">🧠</div>
        <h4>Context-aware chat</h4>
        <p>Ask questions and keep the conversation grounded in your document.</p>
    </div>
    <div class="info-card">
        <div class="info-icon">🔍</div>
        <h4>Smart comparisons</h4>
        <p>Reveal similarities and differences between documents with clarity.</p>
    </div>
</div>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["✨ Summarize", "⚖️ Compare Documents"])

with tab1:
    st.markdown("<div class='section-heading'>Upload a source document and unlock intelligent insights.</div>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload your document", type=["pdf", "docx", "pptx", "txt"])

    if uploaded_file is not None:
        file_size_mb = uploaded_file.size / (1024 * 1024)
        if file_size_mb > MAX_FILE_SIZE_MB:
            st.error(f"❌ File too large ({file_size_mb:.1f}MB). Please upload a file under {MAX_FILE_SIZE_MB}MB.")
            uploaded_file = None
        elif file_size_mb > 5:
            st.warning(f"⚠️ File is {file_size_mb:.1f}MB — processing may take longer.")

    if uploaded_file is not None:
        st.markdown(f"""
        <div class="file-loaded-card">
            <span class="ficon">📄</span>
            <div><b>{uploaded_file.name}</b><span>{uploaded_file.size / 1024:.1f} KB</span></div>
        </div>
        """, unsafe_allow_html=True)

    if uploaded_file is not None and st.session_state.document_text is None:
        doc_text = extract_text(uploaded_file)
        if doc_text:
            st.session_state.document_text = doc_text
            save_history(st.session_state.session_id, st.session_state.chat_history, st.session_state.document_text)
            st.success("✅ Document loaded! Scroll down to chat with it, or click 'Summarize' for a quick summary.")

    st.markdown("<div class='control-panel'>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        language = st.radio("Summary language:", ["English", "Urdu", "Arabic", "French"], horizontal=True)
    with col2:
        summary_length = st.radio("Summary length:", ["Short", "Medium", "Detailed"], horizontal=True)
    st.markdown("</div>", unsafe_allow_html=True)

    if uploaded_file is not None:
        if st.button("Summarize Document"):
            with st.spinner("✍️ Nutshell AI is writing your summary..."):
                summary, tokens_used = summarize_document(st.session_state.document_text, language, summary_length)
            st.session_state.tokens_used_today += tokens_used
            st.session_state.tokens_by_category["summarize"] += tokens_used
            st.session_state.current_summary = summary
            log_analytics("summarize")
            st.rerun()

    if st.session_state.current_summary:
        st.markdown("<div class='control-panel'>", unsafe_allow_html=True)
        st.subheader("Summary")
        st.write(st.session_state.current_summary)
        try:
            pdf_bytes = create_pdf(st.session_state.current_summary)
            st.download_button("📥 Download Summary as PDF", data=pdf_bytes, file_name="summary.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"⚠️ Couldn't prepare the summary PDF: {e}")
        st.markdown("</div>", unsafe_allow_html=True)

with tab2:
    st.markdown("<div class='section-heading'>Upload two documents to uncover nuanced similarities and differences.</div>", unsafe_allow_html=True)
    doc1 = st.file_uploader("Upload first document", type=["pdf", "docx", "pptx", "txt"], key="doc1")
    if doc1 is not None:
        st.markdown(f"""
        <div class="file-loaded-card">
            <span class="ficon">📄</span>
            <div><b>{doc1.name}</b><span>{doc1.size / 1024:.1f} KB</span></div>
        </div>
        """, unsafe_allow_html=True)

    doc2 = st.file_uploader("Upload second document", type=["pdf", "docx", "pptx", "txt"], key="doc2")
    if doc2 is not None:
        st.markdown(f"""
        <div class="file-loaded-card">
            <span class="ficon">📄</span>
            <div><b>{doc2.name}</b><span>{doc2.size / 1024:.1f} KB</span></div>
        </div>
        """, unsafe_allow_html=True)

    compare_language = st.radio("Comparison language:", ["English", "Urdu", "Arabic", "French"], horizontal=True, key="compare_lang")

    if doc1 is not None and doc2 is not None:
        if st.button("Compare Documents"):
            with st.spinner("📊 Nutshell AI is comparing your documents..."):
                text1 = extract_text(doc1)
                text2 = extract_text(doc2)
                if text1 and text2:
                    comparison, tokens_used = compare_documents(text1, text2, compare_language)
                    st.session_state.tokens_used_today += tokens_used
                    st.session_state.tokens_by_category["compare"] += tokens_used
                    st.session_state.compare_document_text = (
                        f"Document 1 ({doc1.name}):\n{text1}\n\n---\n\nDocument 2 ({doc2.name}):\n{text2}"
                    )
                    st.session_state.current_comparison = comparison
                    log_analytics("compare")
                    st.rerun()
                else:
                    st.error("One or both file types are not supported.")

    if st.session_state.get("current_comparison"):
        st.markdown("<div class='control-panel'>", unsafe_allow_html=True)
        st.subheader("Comparison Result")
        st.write(st.session_state.current_comparison)
        st.markdown("</div>", unsafe_allow_html=True)

# ---------------- Persistent chat — collapsed by default, opens on click, works from either tab ----------------

st.divider()

active_chat_context = st.session_state.document_text or st.session_state.compare_document_text

with st.expander("💬 Chat with Nutshell AI", expanded=False):
    st.caption("Ask anything to Nutshell AI.")

    chat_container = st.container()
    with chat_container:
        for msg in st.session_state.chat_history:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])

    user_question = st.chat_input("Type your message here...")
    if user_question:
        st.session_state.chat_history.append({"role": "user", "content": user_question})
        with st.spinner("Nutshell AI is thinking..."):
            answer, tokens_used = chat_with_document(
                active_chat_context,
                user_question,
                st.session_state.chat_history
            )
        st.session_state.tokens_used_today += tokens_used
        st.session_state.tokens_by_category["chat"] += tokens_used
        st.session_state.chat_history.append({"role": "assistant", "content": answer})
        log_analytics("chat")
        save_history(st.session_state.session_id, st.session_state.chat_history, st.session_state.document_text)
        st.rerun()

    if st.session_state.chat_history:
        try:
            chat_export_text = format_chat_for_export(st.session_state.chat_history)
            chat_pdf = create_pdf(chat_export_text)
            st.download_button("📥 Export Chat History as PDF", data=chat_pdf, file_name="chat_history.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"⚠️ Couldn't prepare the chat PDF for download: {e}")