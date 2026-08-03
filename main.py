import json
import requests
import google.generativeai as genai

# ==============================================================================
# CONFIGURATION
# ==============================================================================
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"       # From https://aistudio.google.com/
TELEGRAM_BOT_TOKEN = "YOUR_TELEGRAM_BOT_TOKEN" # From @BotFather
TELEGRAM_CHAT_ID = "YOUR_TELEGRAM_CHAT_ID"     # From @userinfobot

# Initialize Gemini Client
genai.configure(api_key=GEMINI_API_KEY)

def generate_vocab():
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

    model = genai.GenerativeModel('gemini-2.5-flash')
    response = model.generate_content(prompt)
    
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
        print(f"Failed to send message: {req.text}")

if __name__ == "__main__":
    try:
        vocab = generate_vocab()
        send_telegram_message(vocab)
    except Exception as e:
        print(f"Error: {e}")
