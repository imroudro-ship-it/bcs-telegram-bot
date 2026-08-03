import os
import json
import requests
from google import genai

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

def get_best_model(client):
    """Finds an active content generation model associated with the key."""
    try:
        models = client.models.list()
        for m in models:
            # Look for active flash or standard gemini models
            if "generateContent" in getattr(m, "supported_generation_methods", []) or hasattr(m, "name"):
                name = m.name.replace("models/", "")
                if "flash" in name or "gemini" in name:
                    print(f"Selected active model: {name}")
                    return name
    except Exception as e:
        print(f"Model auto-detection skipped: {e}")
    
    # Fallback to standard alias
    return "gemini-2.0-flash"

def generate_vocab():
    client = genai.Client(api_key=GEMINI_API_KEY)
    
    # Auto-detect available active model name
    model_name = get_best_model(client)

    prompt = """
    Act strictly as a professional BCS and Bangladeshi Competitive Job Exam English Mentor.
    Generate 8 high-yield vocabulary words commonly tested in BCS Preliminary/Written, Bank recruitment, and Judicial service exams, focusing on terms found in daily newspapers like The Daily Star or Prothom Alo.

    Return ONLY a raw JSON array of 8 objects (no markdown blocks, no extra text). Each object must have:
    "word": string,
    "pos": string (Noun, Verb, Adjective),
    "level": string (Basic, Intermediate, Advanced),
    "bengali": string (Bangla meaning),
    "synonyms": string (comma-separated),
    "antonyms": string (comma-separated),
    "example": string (Exam-standard sentence)
    """

    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
    )
    
    clean_text = response.text.replace("```json", "").replace("```", "").strip()
    return json.loads(clean_text)

def send_telegram_message(vocab_list):
    message = "<b>📚 DAILY BCS & BANK EXAM VOCABULARY</b>\n"
    message += "<i>Curated from Daily Star & National Dailies</i>\n\n"

    for idx, item in enumerate(vocab_list, 1):
        message += f"<b>{idx}. {item['word']}</b> ({item['pos']}) — <i>{item['level']}</i>\n"
        message += f"• <b>অর্থ:</b> {item['bengali']}\n"
        message += f"• <b>Synonyms:</b> {item['synonyms']}\n"
        message += f"• <b>Antonyms:</b> {item['antonyms']}\n"
        message += f"• <b>Example:</b> <i>\"{item['example']}\"</i>\n\n"

    message += "─────────────────────\n"
    message += "💡 <b>Mentor Tip:</b> Practice using 3 of these words in your written English translation practice today!"

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML"
    }
    
    req = requests.post(url, json=payload)
    if req.status_code == 200:
        print("Telegram message sent successfully!")
    else:
        print(f"Telegram Error: {req.text}")
        raise Exception(f"Telegram API Error: {req.text}")

if __name__ == "__main__":
    if not GEMINI_API_KEY or not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        raise ValueError("Missing required environment secrets.")
    
    vocab = generate_vocab()
    send_telegram_message(vocab)
