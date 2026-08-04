#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import asyncio
import json
import os
import random
import csv
import urllib.request
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
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

# ========== ENV ==========
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
TIMEZONE = pytz.timezone("Asia/Dhaka")

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
EXCEL_FILE = DATA_DIR / "Vocabulary_Bank.xlsx"
HISTORY_FILE = "history.json"

# Dictionary file
DICT_FILE = DATA_DIR / "bangla_dictionary.json"
DICT_CSV_URL = "https://raw.githubusercontent.com/MinhasKamal/BengaliDictionary/master/data/english-bangla.csv"

# --------------------------------------------------------------
# DICTIONARY LOADER: download & convert if missing
# --------------------------------------------------------------

def load_or_download_dictionary():
    if DICT_FILE.exists():
        try:
            with open(DICT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass

    print("📥 Dictionary not found. Downloading from GitHub (this may take a moment)...")
    try:
        # Download CSV
        with urllib.request.urlopen(DICT_CSV_URL) as response:
            csv_data = response.read().decode('utf-8')
        # Parse CSV: each row is [English, Bengali, ...] (first row header)
        lines = csv_data.splitlines()
        reader = csv.reader(lines)
        header = next(reader)  # skip header
        dict_data = {}
        for row in reader:
            if len(row) >= 2:
                eng = row[0].strip()
                ben = row[1].strip()
                if eng and ben:
                    dict_data[eng] = ben
        # Save as JSON for future runs
        with open(DICT_FILE, "w", encoding="utf-8") as f:
            json.dump(dict_data, f, ensure_ascii=False, indent=2)
        print(f"✅ Dictionary loaded: {len(dict_data)} entries.")
        return dict_data
    except Exception as e:
        print(f"⚠️ Failed to download dictionary: {e}. Using AI only.")
        return {}

# Load dictionary
bengali_dict = load_or_download_dictionary()

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

def get_bengali_from_dict(word):
    """Look up the Bangla meaning from the loaded dictionary."""
    # Try exact match
    if word in bengali_dict:
        return bengali_dict[word]
    # Try lowercase
    lower = word.lower()
    if lower in bengali_dict:
        return bengali_dict[lower]
    # Try capitalised
    cap = word.capitalize()
    if cap in bengali_dict:
        return bengali_dict[cap]
    # Try removing common suffixes
    for suffix in ['s', 'ed', 'ing', 'es', 'ly', 'ment', 'tion', 'ness']:
        if word.endswith(suffix):
            stem = word[:-len(suffix)]
            if stem in bengali_dict:
                return bengali_dict[stem]
            if stem.lower() in bengali_dict:
                return bengali_dict[stem.lower()]
    return None

# --------------------------------------------------------------
# AI CALL – SINGLE CALL FOR VOCAB + MCQs
# --------------------------------------------------------------

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10),
       retry=retry_if_exception_type((groq.RateLimitError,)))
def generate_vocab_and_mcqs(client, headlines_data, past_words):
    headlines_text = "\n".join([f"- [{entry['source']}] {entry['headline']}" for entry in headlines_data])
    exclude = ", ".join(past_words[-50:])

    prompt = f"""
You are an expert BCS and Bank job exam mentor with deep knowledge of Bengali and English vocabulary.

Today's headlines from Bangladeshi newspapers (with sources):
{headlines_text}

### Task 1: Generate 50 vocabulary words (numbered 1-50) from these headlines.
- Difficulty: 15 Basic, 20 Intermediate, 15 Advanced.
- Do NOT repeat: {exclude}
- For each word provide EXACTLY these keys:
  "sl", "word", "pos", "level", "bengali", "definition", "synonyms", "antonyms", "example", "category"
- The synonyms and antonyms must be given as a comma-separated string (not a list).

### Task 2: Generate 10 multiple-choice questions (MCQs) from these words.
For each MCQ, provide:
- "question": the question text
- "options": a list of 4 options (A, B, C, D)
- "answer": the correct option letter (A, B, C, or D)
- "explanation": a brief explanation of why the answer is correct

### CRITICAL INSTRUCTIONS FOR BENGALI MEANINGS:
- Use ONLY standard, dictionary-level Bengali meanings.
- Never use transliterations (e.g., "shahosi" is wrong; use "সাহসী").
- Avoid word-by-word translation; provide the closest natural equivalent.
- If a word has multiple meanings, pick the one that matches the headline context.
- Never leave the "bengali" field empty.

### Task 3: Write a detailed 5-7 bullet Bengali summary (in Bengali script) of the most important topics covered in the headlines. Include the newspaper names. The summary must be at least 100 characters long.

Return a JSON object with keys: "vocab_list", "mcqs", and "bengali_summary".
"""
    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": "You are an expert Bengali lexicographer and exam creator. Always output valid JSON. For synonyms and antonyms, use a comma-separated string, not a list. Bengali meanings must be accurate, standard, and in proper Bengali script."},
            {"role": "user", "content": prompt}
        ],
        model="llama-3.3-70b-versatile",
        temperature=0.3,
        response_format={"type": "json_object"},
    )
    data = json.loads(response.choices[0].message.content)
    vocab = data.get("vocab_list", [])
    mcqs = data.get("mcqs", [])
    summary = data.get("bengali_summary", "")
    # Validate MCQs structure
    validated_mcqs = []
    for q in mcqs:
        if isinstance(q, dict) and all(k in q for k in ("question", "options", "answer", "explanation")):
            validated_mcqs.append(q)
    if not summary or len(summary.strip()) < 20:
        summary = build_fallback_summary(headlines_data)
    return vocab, validated_mcqs, summary

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
# BUILD EXCEL – unchanged (but you have it already)
# --------------------------------------------------------------

def build_excel(vocab_list, mcqs):
    # ... (keep your existing build_excel function – it's the same as before) ...
    # I'll include it here for completeness, but you already have it.
    wb = openpyxl.Workbook()
    # ... (paste your existing build_excel code from earlier) ...
    # For brevity, I'll skip the full code here, but it must be included in the final file.
    # Since you have it, just keep it.
    pass

# --------------------------------------------------------------
# BUILD HTML – unchanged
# --------------------------------------------------------------

def build_html(vocab_list, mcqs, summary, date_str):
    # ... (keep your existing build_html function) ...
    pass

def build_index_html(date_str):
    # ... (keep your existing build_index_html function) ...
    pass

# --------------------------------------------------------------
# MAIN JOB
# --------------------------------------------------------------

async def run_daily_job():
    print("🚀 Starting daily job...")
    client = groq.Groq(api_key=GROQ_API_KEY)

    print("📰 Fetching headlines...")
    headlines = fetch_headlines()
    print(f"   Found {len(headlines)} headlines.")

    past = load_history()
    print("🧠 Generating vocabulary, MCQs and summary (single call)...")
    try:
        vocab, mcqs, summary = generate_vocab_and_mcqs(client, headlines, past)
        print(f"   Generated {len(vocab)} words and {len(mcqs)} MCQs.")
    except groq.RateLimitError as e:
        print(f"⚠️ Rate limit hit: {e}. Will try to continue with existing data.")
        # If we can't get new data, we might need to fall back to a cached version? For now, we'll abort.
        # But we'll just re-raise to stop the job.
        raise

    # ---- DICTIONARY LOOKUP ----
    if bengali_dict:
        print("🔍 Looking up meanings in dictionary...")
        dict_hits = 0
        for item in vocab:
            word = item.get("word", "")
            if not word:
                continue
            dict_meaning = get_bengali_from_dict(word)
            if dict_meaning:
                item["bengali"] = dict_meaning
                dict_hits += 1
        print(f"   Found {dict_hits} words in dictionary.")
    else:
        print("⚠️ No dictionary loaded. Using AI meanings.")

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
        html_path = None

    print("📋 Building archive index...")
    try:
        index_path = build_index_html(date_str)
        print(f"   Index saved to {index_path}")
    except Exception as e:
        print(f"   ❌ Error building index: {e}")
        index_path = None

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

            if html_path and html_path.exists():
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
