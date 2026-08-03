import os
import json
import requests
from google import genai
from google.genai import errors

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

def generate_vocab():
    client = genai.Client(api_key=GEMINI_API_KEY)
    
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

    # Primary and fallback models for Google GenAI SDK
    models_to_try = ["gemini-2.0-flash", "gemini-1.5-flash"]
    
    response = None
    for model_name in models_to_try:
        try:
            print(f"Attempting generation with model: {model_name}...")
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
            )
            print(f"Successfully generated response using {model_name}!")
            break
        except errors.APIError as e:
            print(f"Warning: {model_name} failed with error: {e}")
            continue

    if not response or not response.text:
        raise Exception("Failed to generate content using all available Gemini models.")
    
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
