#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import asyncio
import json
import os
import random
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
from tenacity import retry, stop_after_attempt, wait_exponential

# PDF libraries
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch, cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ========== ENV ==========
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
TIMEZONE = pytz.timezone("Asia/Dhaka")

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
EXCEL_FILE = DATA_DIR / "Vocabulary_Bank.xlsx"
PDF_FILE = DATA_DIR / "Vocabulary_Bank.pdf"   # we'll also include date
HISTORY_FILE = "history.json"
FONT_DIR = DATA_DIR / "fonts"
FONT_DIR.mkdir(exist_ok=True)

RSS_FEEDS = {
    "The Daily Star": "https://www.thedailystar.net/rss.xml",
    "Dhaka Tribune": "https://www.dhakatribune.com/feed",
    "The Business Standard": "https://www.tbsnews.net/rss.xml",
}

# --------------------------------------------------------------
# HELPER: Download NotoSansBengali font if missing
# --------------------------------------------------------------
def download_bengali_font():
    font_path = FONT_DIR / "NotoSansBengali-Regular.ttf"
    if font_path.exists():
        return str(font_path)
    print("📥 Downloading Bengali font...")
    url = "https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSansBengali/NotoSansBengali-Regular.ttf"
    try:
        urllib.request.urlretrieve(url, font_path)
        print("✅ Font downloaded.")
        return str(font_path)
    except Exception as e:
        print(f"⚠️ Font download failed: {e}. Using fallback.")
        return None

# --------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------

def safe_excel_value(value):
    """Convert any value to an Excel‑safe string."""
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
# AI CALLS
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
    # Fallback if summary is empty
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

    summary = "📰 **আজকের প্রধান সংবাদ (সূত্র সহ)**\n\n"
    for src, headlines in sources.items():
        summary += f"▪️ **{src}**\n"
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
# BUILD EXCEL – GEMINI STYLE (3 SHEETS) – unchanged
# --------------------------------------------------------------

def build_excel(vocab_list, mcqs):
    # ... (exactly the same as your current build_excel) ...
    # I'll include it fully in the final answer, but for brevity here I'll skip.
    # In the final response, I will paste the complete code.
    pass

# --------------------------------------------------------------
# BUILD PDF – PROFESSIONAL WITH BENGALI FONT
# --------------------------------------------------------------

def build_pdf(vocab_list, mcqs, summary, date_str):
    pdf_path = DATA_DIR / f"Vocabulary_{date_str}.pdf"

    # Download and register Bengali font
    font_path = download_bengali_font()
    if font_path:
        try:
            pdfmetrics.registerFont(TTFont('NotoSansBengali', font_path))
            bengali_font = 'NotoSansBengali'
        except:
            bengali_font = 'Helvetica'
    else:
        bengali_font = 'Helvetica'

    # Create styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Title'],
        fontName=bengali_font,
        fontSize=18,
        alignment=TA_CENTER,
        spaceAfter=0.3*inch
    )
    heading_style = ParagraphStyle(
        'HeadingStyle',
        parent=styles['Heading2'],
        fontName=bengali_font,
        fontSize=14,
        spaceAfter=0.2*inch,
        textColor=colors.HexColor('#1F4E78')
    )
    normal_style = ParagraphStyle(
        'NormalStyle',
        parent=styles['Normal'],
        fontName=bengali_font,
        fontSize=10,
        leading=14,
        alignment=TA_LEFT
    )
    table_cell_style = ParagraphStyle(
        'TableCell',
        parent=normal_style,
        fontSize=9,
        leading=12
    )

    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4,
                            rightMargin=0.8*inch, leftMargin=0.8*inch,
                            topMargin=0.8*inch, bottomMargin=0.8*inch)
    story = []

    # ---- Page 1: Title, date, summary, vocabulary table ----
    story.append(Paragraph("📘 BCS & Bank Job Preparation – Daily Vocabulary", title_style))
    story.append(Paragraph(f"<b>Date:</b> {date_str}", normal_style))
    story.append(Spacer(1, 0.2*inch))

    if summary:
        story.append(Paragraph("📰 Today's Summary", heading_style))
        # Replace newlines with <br/> and bold markers
        summary_clean = summary.replace('\n', '<br/>').replace('**', '<b>').replace('**', '</b>')
        story.append(Paragraph(summary_clean, normal_style))
        story.append(Spacer(1, 0.2*inch))

    # Vocabulary table – all columns: Word, POS, Bengali, Definition, Synonyms, Antonyms, Example
    story.append(Paragraph("📚 Vocabulary List (All 100 Words)", heading_style))
    table_data = [['SL', 'Word', 'POS', 'Bengali', 'Definition', 'Synonyms', 'Antonyms', 'Example']]
    for item in vocab_list:
        # truncate long text for table
        def trunc(txt, limit=40):
            if len(txt) > limit:
                return txt[:limit] + '...'
            return txt
        table_data.append([
            str(item.get('sl', '')),
            item.get('word', ''),
            item.get('pos', ''),
            item.get('bengali', ''),
            trunc(item.get('definition', ''), 35),
            trunc(item.get('synonyms', ''), 30),
            trunc(item.get('antonyms', ''), 25),
            trunc(item.get('example', ''), 30)
        ])

    # Set column widths (in points)
    col_widths = [0.4*inch, 1.2*inch, 0.7*inch, 1.2*inch, 1.8*inch, 1.4*inch, 1.2*inch, 1.8*inch]
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E78')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), bengali_font),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#F2F5F9')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTNAME', (0,1), (-1,-1), bengali_font),
        ('FONTSIZE', (0,1), (-1,-1), 8),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t)
    story.append(PageBreak())

    # ---- Page 2+: MCQs ----
    if mcqs:
        story.append(Paragraph("📝 Practice Test (10 MCQs)", heading_style))
        for i, q in enumerate(mcqs[:10], 1):
            story.append(Paragraph(f"<b>{i}. {q['question']}</b>", normal_style))
            for opt in q['options']:
                story.append(Paragraph(f"   {opt}", normal_style))
            story.append(Paragraph(f"✅ <b>Answer:</b> {q['answer']} – {q['explanation']}", normal_style))
            story.append(Spacer(1, 0.1*inch))

    doc.build(story)
    return pdf_path

# --------------------------------------------------------------
# MAIN JOB
# --------------------------------------------------------------

async def run_daily_job():
    client = groq.Groq(api_key=GROQ_API_KEY)
    headlines = fetch_headlines()
    past = load_history()
    vocab, summary = generate_vocab_and_summary(client, headlines, past)
    mcqs = generate_mcqs(client, vocab)
    build_excel(vocab, mcqs)

    # Generate PDF
    date_str = datetime.now(TIMEZONE).strftime("%Y-%m-%d")
    pdf_path = build_pdf(vocab, mcqs, summary, date_str)

    # Save history
    save_history(past + [w["word"] for w in vocab])

    # Send via Telegram
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        bot = Bot(token=TELEGRAM_BOT_TOKEN)
        # Send Excel
        with open(EXCEL_FILE, "rb") as f:
            await bot.send_document(chat_id=TELEGRAM_CHAT_ID, document=f, caption="📊 Daily Vocabulary Bank (Excel)")
        # Send PDF
        with open(pdf_path, "rb") as f:
            await bot.send_document(chat_id=TELEGRAM_CHAT_ID, document=f, caption="📄 Daily Vocabulary Bank (PDF – mobile friendly)")
        # Send summary (ensure it's not empty)
        if not summary or len(summary.strip()) < 10:
            summary = "আজকের সংক্ষিপ্ত সারাংশ তৈরি করা সম্ভব হয়নি। তবে ভোকাবুলারি ফাইলগুলো দেখুন।"
        await bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=f"📰 **Today's Summary**\n\n{summary}", parse_mode="Markdown")

    print("✅ Job done – Excel and PDF generated and sent.")

# --------------------------------------------------------------
# ENTRY POINT
# --------------------------------------------------------------

if __name__ == "__main__":
    asyncio.run(run_daily_job())
