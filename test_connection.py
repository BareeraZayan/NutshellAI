from pymongo import MongoClient
from dotenv import load_dotenv
import os

load_dotenv()

uri = os.getenv("MONGO_URI")

# Password ko mask karke print karo, taake dekh sakein URI kaisi load hui
if uri:
    parts = uri.split("@")
    print("URI Start (masked):", parts[0][:20] + "...MASKED...")
    print("URI End:", "@" + parts[1] if len(parts) > 1 else "No @ found - PROBLEM!")
else:
    print("MONGO_URI is None - .env file load nahi hui ya key galat likhi hai!")

print("\n--- Testing Connection ---")
try:
    client = MongoClient(uri, serverSelectionTimeoutMS=5000)
    client.admin.command('ping')
    print("✅ SUCCESS! Connection working!")
except Exception as e:
    print("❌ FAILED:", e)