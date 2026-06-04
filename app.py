from pymongo import MongoClient
from groq import Groq
from dotenv import load_dotenv
from PyPDF2 import PdfReader
from docx import Document
from pptx import Presentation
import os

# ---------------- Extraction Functions ----------------

def extract_text_from_pdf(file):
    reader = PdfReader(file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
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

load_dotenv()

client = Groq(api_key=os.getenv("GROQ_API_KEY"))
GROQ_MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-120b")

LANG_MAP = {"English": "in English", "Urdu": "in Urdu", "Arabic": "in Arabic", "French": "in French"}

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

mongo_client = MongoClient(os.getenv("MONGO_URI"))
db = mongo_client["summarizer_db"]
collection = db["chat_sessions"]

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
