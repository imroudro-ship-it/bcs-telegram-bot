import os
import json
import requests
from groq import Groq

GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

def generate_vocab():
    client = Groq(api_key=GROQ_API_KEY)
    
    prompt = """
    Act strictly as a professional BCS and Bangladeshi Competitive Job Exam English Mentor.
    Generate 8 high-yield vocabulary words commonly tested in BCS Preliminary/Written, Bank recruitment, and Judicial service exams, focusing on terms found in daily newspapers like The Daily Star or Prothom Alo.

    Return ONLY a raw JSON array of 8 objects with no markdown formatting or extra commentary. 
    Keys for each object:
    "word": string,
    "pos": string (Noun, Verb, Adjective),
    "level": string (Basic, Intermediate, Advanced),
    "bengali": string (Bangla meaning),
    "synonyms": string (comma-separated),
    "antonyms": string (comma-separated),
    "example": string (Exam-standard sentence)
    """

    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": "You are a JSON-only generator for Bangladeshi competitive job exams."},
            {"role": "user", "content": prompt}
        ],
        model="llama-3.3-70b-versatile",
        temperature=0.5,
    )
    
    raw_text = response.choices[0].message.content
    clean_text = raw_text.replace("```json", "").replace("```", "").strip()
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
        print("Telegram message sent successfully via Groq AI!")
    else:
        print(f"Telegram Error: {req.text}")
        raise Exception(f"Telegram API Error: {req.text}")

if __name__ == "__main__":
    if not GROQ_API_KEY or not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        raise ValueError("Missing required environment secrets.")
    
    vocab = generate_vocab()
    send_telegram_message(vocab)
