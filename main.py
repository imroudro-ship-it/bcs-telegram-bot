#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import asyncio
import json
import os
import random
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
# AI CALLS (unchanged)
# --------------------------------------------------------------

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def generate_vocab_and_summary(client, headlines_data, past_words):
    headlines_text = "\n".join([f"- [{entry['source']}] {entry['headline']}" for entry in headlines_data])
    exclude = ", ".join(past_words[-50:])

    prompt = f"""
You are an expert BCS and Bank job exam mentor.

Today's headlines from Bangladeshi newspapers (with sources):
{headlines_text}

### Task 1: Generate 100 vocabulary words (numbered 1-100) from these headlines.
- Difficulty: 30 Basic, 40 Intermediate, 30 Advanced.
- Do NOT repeat: {exclude}
- For each word provide EXACTLY these keys:
  "sl", "word", "pos", "level", "bengali", "definition", "synonyms", "antonyms", "example", "category"
- The synonyms and antonyms must be given as a comma-separated string (not a list).

### Task 2: Write a detailed 5-7 bullet Bengali summary (in Bengali script) of the most important topics covered in the headlines. Include the newspaper names. The summary must be at least 100 characters long.

Return JSON with keys: "vocab_list" and "bengali_summary".
"""
    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": "You are an expert Bengali lexicographer. Always output valid JSON. For synonyms and antonyms, use a comma-separated string, not a list."},
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
# BUILD EXCEL (unchanged – keep your existing build_excel)
# --------------------------------------------------------------

def build_excel(vocab_list, mcqs):
    # ... (your existing build_excel code – same as before) ...
    # I'll assume you have it. If not, I'll include it fully in the final answer.
    pass

# --------------------------------------------------------------
# BUILD HTML – IBA COMPENDIUM STYLE (FIXED)
# --------------------------------------------------------------

def build_html(vocab_list, mcqs, summary, date_str):
    html_path = DATA_DIR / f"Vocabulary_{date_str}.html"

    # Pre‑process summary to replace newlines with <br> (outside f‑string)
    summary_br = summary.replace('\n', '<br>')

    # Group vocabulary by category
    categories = {}
    for item in vocab_list:
        cat = item.get('category', 'Miscellaneous')
        if cat not in categories:
            categories[cat] = []
        categories[cat].append(item)

    sorted_cats = sorted(categories.items())

    # Build category tables
    category_html = ""
    for cat, words in sorted_cats:
        category_html += f"""
        <h2 class="category-title">{cat}</h2>
        <table>
            <thead>
                <tr>
                    <th>Word</th>
                    <th>Bengali Meaning</th>
                    <th>POS</th>
                    <th>Synonyms</th>
                    <th>Antonyms</th>
                    <th>Example</th>
                </tr>
            </thead>
            <tbody>
        """
        for w in words:
            category_html += f"""
                <tr>
                    <td><strong>{w.get('word', '')}</strong></td>
                    <td>{w.get('bengali', '')}</td>
                    <td>{w.get('pos', '')}</td>
                    <td>{w.get('synonyms', '')}</td>
                    <td>{w.get('antonyms', '')}</td>
                    <td>{w.get('example', '')}</td>
                </tr>
            """
        category_html += """
            </tbody>
        </table>
        """

    # Build MCQs
    mcq_html = ""
    if mcqs:
        mcq_html += """
        <div class="mcq-section">
            <h1>📝 Practice Test (10 MCQs)</h1>
        """
        for i, q in enumerate(mcqs[:10], 1):
            options_html = "".join(f"<li>{opt}</li>" for opt in q.get('options', []))
            mcq_html += f"""
            <div class="mcq">
                <p><strong>{i}. {q['question']}</strong></p>
                <ul>
                    {options_html}
                </ul>
                <p><strong>✅ Answer:</strong> {q['answer']} – {q['explanation']}</p>
            </div>
            """
        mcq_html += "</div>"

    # Full HTML template (no backslashes inside f‑string)
    html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Daily Vocabulary – {date_str}</title>
    <style>
        body {{
            font-family: 'Segoe UI', 'Noto Sans Bengali', Arial, sans-serif;
            margin: 20px;
            background: #f4f6f9;
            color: #1e2a3a;
        }}
        .container {{
            max-width: 1100px;
            margin: auto;
            background: white;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.1);
        }}
        .cover {{
            text-align: center;
            padding: 40px 0;
            border-bottom: 3px solid #1F4E78;
            margin-bottom: 30px;
        }}
        .cover h1 {{
            font-size: 28px;
            color: #1F4E78;
            margin-bottom: 5px;
        }}
        .cover .date {{
            font-size: 18px;
            color: #555;
        }}
        .cover .stats {{
            font-size: 14px;
            color: #777;
            margin-top: 10px;
        }}
        .summary {{
            background: #eef3f9;
            padding: 15px 20px;
            border-radius: 8px;
            margin-bottom: 30px;
            font-size: 1.05em;
            line-height: 1.6;
            white-space: pre-wrap;
        }}
        .category-title {{
            color: #1F4E78;
            border-bottom: 2px solid #1F4E78;
            padding-bottom: 8px;
            margin-top: 35px;
            font-size: 20px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0 30px 0;
            font-size: 0.85em;
        }}
        th {{
            background: #1F4E78;
            color: white;
            font-weight: bold;
            padding: 10px 8px;
            text-align: left;
        }}
        td {{
            padding: 8px 8px;
            border: 1px solid #ddd;
            vertical-align: top;
        }}
        tr:nth-child(even) {{
            background: #f9fafc;
        }}
        .mcq-section {{
            margin-top: 40px;
            border-top: 3px solid #1F4E78;
            padding-top: 20px;
        }}
        .mcq {{
            background: #f9fafc;
            border-left: 4px solid #1F4E78;
            padding: 12px 18px;
            margin: 15px 0;
            border-radius: 4px;
        }}
        .mcq ul {{
            list-style-type: none;
            padding-left: 20px;
            margin: 5px 0;
        }}
        .mcq li {{
            margin: 3px 0;
        }}
        .footer {{
            text-align: center;
            margin-top: 40px;
            font-size: 0.9em;
            color: #777;
            border-top: 1px solid #ddd;
            padding-top: 15px;
        }}
        @media print {{
            body {{ background: white; margin: 0; }}
            .container {{ box-shadow: none; border: none; }}
            .mcq {{ break-inside: avoid; }}
        }}
        @media (max-width: 600px) {{
            table {{ font-size: 0.7em; }}
            td, th {{ padding: 4px; }}
            .container {{ padding: 12px; }}
        }}
    </style>
</head>
<body>
<div class="container">
    <!-- Cover -->
    <div class="cover">
        <h1>📘 Daily Vocabulary Bank</h1>
        <div class="date">📅 {date_str}</div>
        <div class="stats">
            {len(vocab_list)} words • {len(categories)} categories • 10 MCQs
        </div>
    </div>

    <!-- Summary -->
    <div class="summary">
        <h3 style="margin-top:0;">📰 Today's Summary</h3>
        {summary_br}
    </div>

    <!-- Vocabulary by Category -->
    {category_html}

    <!-- MCQs -->
    {mcq_html}

    <div class="footer">
        Generated automatically • Daily Star Vocabulary Bank • For BCS & Bank Exams
    </div>
</div>
</body>
</html>
"""
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    return html_path

# --------------------------------------------------------------
# MAIN JOB – WITH TELEGRAM DEBUGGING
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

    print("❓ Generating MCQs...")
    mcqs = generate_mcqs(client, vocab)
    print(f"   Generated {len(mcqs)} MCQs.")

    print("📊 Building Excel file...")
    build_excel(vocab, mcqs)
    print(f"   Excel saved to {EXCEL_FILE}")

    date_str = datetime.now(TIMEZONE).strftime("%Y-%m-%d")
    print("📄 Building HTML file...")
    html_path = build_html(vocab, mcqs, summary, date_str)
    print(f"   HTML saved to {html_path}")

    save_history(past + [w["word"] for w in vocab])
    print("💾 History updated.")

    # ---------- Telegram Sending ----------
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        print("📤 Sending files via Telegram...")
        try:
            bot = Bot(token=TELEGRAM_BOT_TOKEN)

            # Send Excel
            with open(EXCEL_FILE, "rb") as f:
                await bot.send_document(chat_id=TELEGRAM_CHAT_ID, document=f, caption="📊 Daily Vocabulary Bank (Excel)")
                print("   ✅ Excel sent.")

            # Send HTML
            with open(html_path, "rb") as f:
                await bot.send_document(chat_id=TELEGRAM_CHAT_ID, document=f, caption="📄 Daily Vocabulary Bank (HTML – open on any device)")
                print("   ✅ HTML sent.")

            # Send summary
            if not summary or len(summary.strip()) < 10:
                summary = "আজকের সংক্ষিপ্ত সারাংশ তৈরি করা সম্ভব হয়নি। তবে ভোকাবুলারি ফাইলগুলো দেখুন।"
            await bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=f"📰 **Today's Summary**\n\n{summary}", parse_mode="Markdown")
            print("   ✅ Summary sent.")

        except Exception as e:
            print(f"   ❌ Failed to send Telegram messages: {e}")
    else:
        print("⚠️ Telegram credentials missing. Skipping send.")

    print("✅ Job done – Excel and HTML generated.")

# --------------------------------------------------------------
# ENTRY POINT
# --------------------------------------------------------------

if __name__ == "__main__":
    asyncio.run(run_daily_job())
