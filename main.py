#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import asyncio
import json
import os
import random
import re
from datetime import datetime
from pathlib import Path

import feedparser
import groq
import openpyxl
import pytz
import requests
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from telegram import Bot
from tenacity import retry, stop_after_attempt, wait_exponential

# ========== ENV ==========
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
TIMEZONE = pytz.timezone("Asia/Dhaka")

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
EXCEL_FILE = DATA_DIR / "Vocabulary_Bank.xlsx"
HISTORY_FILE = "history.json"

RSS_FEEDS = {
    "The Daily Star": "https://www.thedailystar.net/rss.xml",
    "Dhaka Tribune": "https://www.dhakatribune.com/feed",
    "The Business Standard": "https://www.tbsnews.net/rss.xml",
}

# --------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------

def safe_excel_value(value):
    if value is None:
        return ""
    if isinstance(value, list):
        return ", ".join(str(v) for v in value)
    if isinstance(value, dict):
        return str(value)
    return str(value)

def fetch_headlines():
    all_entries = []
    for source, url in RSS_FEEDS.items():
        try:
            feed = feedparser.parse(url)
            for entry in feed.entries[:5]:
                all_entries.append({"headline": entry.title, "source": source})
        except:
            pass
    seen = set()
    unique = []
    for entry in all_entries:
        if entry["headline"] not in seen:
            seen.add(entry["headline"])
            unique.append(entry)
    return unique[:25]

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r") as f:
            return json.load(f).get("words", [])
    return []

def save_history(words):
    with open(HISTORY_FILE, "w") as f:
        json.dump({"words": words}, f)

# --------------------------------------------------------------
# SUSPICIOUS MEANING DETECTION
# --------------------------------------------------------------

def is_bengali_meaning_suspicious(meaning):
    """Return True if the meaning looks incorrect or incomplete."""
    if not meaning or len(meaning.strip()) < 2:
        return True
    # If it contains any English letter (a-z or A-Z)
    if re.search(r'[a-zA-Z]', meaning):
        return True
    # Generic placeholders
    if meaning.strip() in ["অর্থ", "নেই", "অজানা", "শব্দ"]:
        return True
    return False

# --------------------------------------------------------------
# AI CALLS
# --------------------------------------------------------------

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def generate_vocab_and_summary(client, headlines_data, past_words):
    headlines_text = "\n".join([f"- [{entry['source']}] {entry['headline']}" for entry in headlines_data])
    exclude = ", ".join(past_words[-50:])

    prompt = f"""
You are an expert BCS and Bank job exam mentor with deep knowledge of Bengali and English vocabulary.

Today's headlines from Bangladeshi newspapers (with sources):
{headlines_text}

### Task 1: Generate 100 vocabulary words (numbered 1-100) from these headlines.
- Difficulty: 30 Basic, 40 Intermediate, 30 Advanced.
- Do NOT repeat: {exclude}
- For each word provide EXACTLY these keys:
  "sl", "word", "pos", "level", "bengali", "definition", "synonyms", "antonyms", "example", "category"
- The synonyms and antonyms must be given as a comma-separated string (not a list).

### CRITICAL INSTRUCTIONS FOR BENGALI MEANINGS:
- Use ONLY standard, dictionary-level Bengali meanings.
- Never use transliterations (e.g., "shahosi" is wrong; use "সাহসী").
- Avoid word-by-word translation; provide the closest natural equivalent.
- If a word has multiple meanings, pick the one that matches the headline context.
- Never leave the "bengali" field empty.

### Task 2: Write a detailed 5-7 bullet Bengali summary (in Bengali script) of the most important topics covered in the headlines. Include the newspaper names. The summary must be at least 100 characters long.

Return JSON with keys: "vocab_list" and "bengali_summary".
"""
    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": "You are an expert Bengali lexicographer. Always output valid JSON. For synonyms and antonyms, use a comma-separated string, not a list. Bengali meanings must be accurate, standard, and in proper Bengali script."},
            {"role": "user", "content": prompt}
        ],
        model="llama-3.3-70b-versatile",
        temperature=0.3,
        response_format={"type": "json_object"},
    )
    data = json.loads(response.choices[0].message.content)
    vocab = data.get("vocab_list", [])
    summary = data.get("bengali_summary", "")
    if not summary or len(summary.strip()) < 20:
        summary = build_fallback_summary(headlines_data)
    return vocab, summary

def build_fallback_summary(headlines_data):
    sources = {}
    for entry in headlines_data:
        src = entry["source"]
        if src not in sources:
            sources[src] = []
        sources[src].append(entry["headline"])

    summary = "📰 আজকের প্রধান সংবাদ (সূত্র সহ)\n\n"
    for src, headlines in sources.items():
        summary += f"▪️ {src}\n"
        for h in headlines[:3]:
            summary += f"   - {h}\n"
        summary += "\n"
    summary += "🔹 অন্যান্য সংবাদপত্র থেকেও গুরুত্বপূর্ণ শব্দ সংগ্রহ করা হয়েছে।"
    return summary

# --------------------------------------------------------------
# FALLBACK: Correct suspicious Bengali meanings
# --------------------------------------------------------------

def fix_suspicious_meanings(client, vocab_list, headlines_data):
    """Re-generate Bengali meanings for words that look suspicious."""
    # Build a dict of headlines for context
    headline_map = {}
    for entry in headlines_data:
        # We'll use the first headline that contains the word (approximate)
        # For simplicity, we'll store all headlines
        pass

    for item in vocab_list:
        word = item.get("word", "")
        bengali = item.get("bengali", "")
        if not word:
            continue
        if is_bengali_meaning_suspicious(bengali):
            print(f"   ⚠️ Suspicious meaning for '{word}': '{bengali}' – re‑generating...")
            # Call AI to get a fresh translation for this word
            try:
                corrected = get_corrected_bengali(client, word, headlines_data)
                if corrected:
                    item["bengali"] = corrected
                    print(f"      ✅ Corrected to '{corrected}'")
                else:
                    print(f"      ❌ Fallback failed – keeping original")
            except Exception as e:
                print(f"      ❌ Error in fallback: {e}")
    return vocab_list

def get_corrected_bengali(client, word, headlines_data):
    """Ask the AI specifically for a Bengali meaning of a single word with context."""
    # Gather context from headlines that mention the word (or just use the first few)
    context = " ".join([entry["headline"] for entry in headlines_data[:5]])
    prompt = f"""
You are a Bengali-English dictionary expert.

The English word is "{word}".
It appears in news headlines like:
{context}

Please provide the most accurate, standard Bengali translation (in Bengali script) for "{word}" in this context.
Return ONLY the Bengali meaning, nothing else. No explanation, no quotes.
"""
    try:
        response = client.chat.completions.create(
            messages=[
                {"role": "system", "content": "You are a Bengali dictionary expert. Provide only the Bengali meaning in Bengali script."},
                {"role": "user", "content": prompt}
            ],
            model="llama-3.3-70b-versatile",
            temperature=0.2,
            max_tokens=20,
        )
        meaning = response.choices[0].message.content.strip()
        # Remove extra punctuation or quotes
        meaning = meaning.strip('"').strip("'").strip()
        if meaning and not is_bengali_meaning_suspicious(meaning):
            return meaning
        else:
            return None
    except Exception as e:
        print(f"      Fallback API error: {e}")
        return None

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def generate_mcqs(client, vocab_list):
    selected = random.sample(vocab_list, min(10, len(vocab_list)))
    word_list = [w["word"] for w in selected]

    prompt = f"""
Generate 10 multiple-choice questions (MCQs) from these vocabulary words: {', '.join(word_list)}.
For each question, provide:
- "question": the question text
- "options": a list of 4 options (A, B, C, D)
- "answer": the correct option letter (A, B, C, or D)
- "explanation": a brief explanation of why the answer is correct

Return a JSON array of objects. The array must contain exactly 10 objects.
Each object must have keys: question, options, answer, explanation.
"""
    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": "You are an exam creator. Always output valid JSON."},
            {"role": "user", "content": prompt}
        ],
        model="llama-3.3-70b-versatile",
        temperature=0.4,
        response_format={"type": "json_object"},
    )
    try:
        raw = json.loads(response.choices[0].message.content)
        if isinstance(raw, dict):
            if "mcqs" in raw:
                mcqs = raw["mcqs"]
            elif "questions" in raw:
                mcqs = raw["questions"]
            else:
                for v in raw.values():
                    if isinstance(v, list):
                        mcqs = v
                        break
                else:
                    mcqs = []
        elif isinstance(raw, list):
            mcqs = raw
        else:
            mcqs = []
        validated = []
        for q in mcqs:
            if isinstance(q, dict) and all(k in q for k in ("question", "options", "answer", "explanation")):
                validated.append(q)
        return validated
    except Exception as e:
        print(f"⚠️ MCQs parsing error: {e}")
        return []

# --------------------------------------------------------------
# BUILD EXCEL (unchanged – reuse your existing build_excel)
# --------------------------------------------------------------

def build_excel(vocab_list, mcqs):
    # ... (keep your existing build_excel function exactly as it is) ...
    # For brevity, I'll not paste it here – you already have it.
    pass

# --------------------------------------------------------------
# BUILD HTML (unchanged – reuse your existing build_html)
# --------------------------------------------------------------

def build_html(vocab_list, mcqs, summary, date_str):
    # ... (keep your existing build_html function exactly as it is) ...
    # For brevity, I'll not paste it here – you already have it.
    pass

# --------------------------------------------------------------
# BUILD INDEX PAGE (unchanged)
# --------------------------------------------------------------

def build_index_html(date_str):
    # ... (keep your existing build_index_html function exactly as it is) ...
    # For brevity, I'll not paste it here – you already have it.
    pass

# --------------------------------------------------------------
# MAIN JOB – WITH FALLBACK
# --------------------------------------------------------------

async def run_daily_job():
    print("🚀 Starting daily job...")
    client = groq.Groq(api_key=GROQ_API_KEY)

    print("📰 Fetching headlines...")
    headlines = fetch_headlines()
    print(f"   Found {len(headlines)} headlines.")

    past = load_history()
    print("🧠 Generating vocabulary and summary...")
    vocab, summary = generate_vocab_and_summary(client, headlines, past)
    print(f"   Generated {len(vocab)} words.")

    # ---- FALLBACK: Fix suspicious Bengali meanings ----
    print("🔍 Checking and correcting Bengali meanings...")
    vocab = fix_suspicious_meanings(client, vocab, headlines)

    print("❓ Generating MCQs...")
    mcqs = generate_mcqs(client, vocab)
    print(f"   Generated {len(mcqs)} MCQs.")

    print("📊 Building Excel file...")
    try:
        build_excel(vocab, mcqs)
        print(f"   Excel saved to {EXCEL_FILE}")
    except Exception as e:
        print(f"   ❌ Error building Excel: {e}")

    date_str = datetime.now(TIMEZONE).strftime("%Y-%m-%d")
    print("📄 Building HTML file...")
    try:
        html_path = build_html(vocab, mcqs, summary, date_str)
        print(f"   HTML saved to {html_path}")
    except Exception as e:
        print(f"   ❌ Error building HTML: {e}")

    print("📋 Building archive index...")
    try:
        index_path = build_index_html(date_str)
        print(f"   Index saved to {index_path}")
    except Exception as e:
        print(f"   ❌ Error building index: {e}")

    save_history(past + [w["word"] for w in vocab])
    print("💾 History updated.")

    # ---------- Telegram Sending ----------
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        print("📤 Sending files via Telegram...")
        try:
            bot = Bot(token=TELEGRAM_BOT_TOKEN)
            if EXCEL_FILE.exists():
                with open(EXCEL_FILE, "rb") as f:
                    await bot.send_document(chat_id=TELEGRAM_CHAT_ID, document=f, caption="📊 Daily Vocabulary Bank (Excel)")
                print("   ✅ Excel sent.")
            else:
                print("   ⚠️ Excel file not found – skipping.")
            if html_path.exists():
                with open(html_path, "rb") as f:
                    await bot.send_document(chat_id=TELEGRAM_CHAT_ID, document=f, caption="📄 Daily Vocabulary Bank (HTML – open on any device)")
                print("   ✅ HTML sent.")
            else:
                print("   ⚠️ HTML file not found – skipping.")
            if not summary or len(summary.strip()) < 10:
                summary = "আজকের সংক্ষিপ্ত সারাংশ তৈরি করা সম্ভব হয়নি। তবে ভোকাবুলারি ফাইলগুলো দেখুন।"
            await bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=f"📰 **Today's Summary**\n\n{summary}", parse_mode="Markdown")
            print("   ✅ Summary sent.")
        except Exception as e:
            print(f"   ❌ Failed to send Telegram messages: {e}")
    else:
        print("⚠️ Telegram credentials missing. Skipping send.")

    print("✅ Job done – Excel, HTML and index generated.")

if __name__ == "__main__":
    asyncio.run(run_daily_job())
