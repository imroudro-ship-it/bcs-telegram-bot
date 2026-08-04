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

# PDF libraries
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

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

RSS_FEEDS = {
    "The Daily Star": "https://www.thedailystar.net/rss.xml",
    "Dhaka Tribune": "https://www.dhakatribune.com/feed",
    "The Business Standard": "https://www.tbsnews.net/rss.xml",
}

# --------------------------------------------------------------
# HELPER: Try to register a Bengali font if available
# --------------------------------------------------------------
def register_bengali_font():
    try:
        # Look for a common Bengali font on the system
        font_paths = [
            "/usr/share/fonts/truetype/noto/NotoSansBengali-Regular.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",  # fallback
            "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",  # macOS
            "C:/Windows/Fonts/ArialUnicode.ttf",  # Windows
        ]
        for path in font_paths:
            if os.path.exists(path):
                pdfmetrics.registerFont(TTFont('BengaliFont', path))
                return 'BengaliFont'
    except:
        pass
    return 'Helvetica'  # fallback – may not show Bengali perfectly

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
# BUILD EXCEL – GEMINI STYLE (3 SHEETS)
# --------------------------------------------------------------

def build_excel(vocab_list, mcqs):
    wb = openpyxl.Workbook()

    # ---------- Sheet 1: Vocabulary ----------
    ws = wb.active
    ws.title = "Daily Star Vocabulary"

    # Header styling
    navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    light_navy = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    even = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))

    # Title rows (merged)
    ws.merge_cells("A1:J1")
    ws["A1"] = "THE DAILY STAR - DAILY VOCABULARY BANK (BCS & JOB PREP SPECIAL)"
    ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = navy
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:J2")
    ws["A2"] = "Curated from Today's Headlines, Editorials & Reports | Includes Meanings, Parts of Speech, Bengali Translations & Exam Examples"
    ws["A2"].font = Font(size=10, italic=True, color="FFFFFF")
    ws["A2"].fill = light_navy
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["SL No", "Word / Idiom", "Part of Speech", "Difficulty Level",
               "Bengali Meaning (বাংলা অর্থ)", "English Definition",
               "Synonyms", "Antonyms", "Contextual Example (Daily Star)", "Category / Topic"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col)
        cell.value = h
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = navy
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Write vocabulary rows
    for idx, item in enumerate(vocab_list, 5):
        ws.row_dimensions[idx].height = 26
        row_data = [
            idx - 4,  # SL No
            safe_excel_value(item.get("word", "")),
            safe_excel_value(item.get("pos", "")),
            safe_excel_value(item.get("level", "")),
            safe_excel_value(item.get("bengali", "")),
            safe_excel_value(item.get("definition", "")),
            safe_excel_value(item.get("synonyms", "")),
            safe_excel_value(item.get("antonyms", "")),
            safe_excel_value(item.get("example", "")),
            safe_excel_value(item.get("category", ""))
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=idx, column=col)
            cell.value = val
            cell.border = border
            cell.fill = even if idx % 2 == 0 else white
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if col in [1, 3, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Color‑code difficulty
            if col == 4:
                level = str(val).strip().capitalize()
                if level == "Basic":
                    cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
                    cell.font = Font(bold=True, color="2E7D32")
                elif level == "Intermediate":
                    cell.fill = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
                    cell.font = Font(bold=True, color="F57F17")
                elif level == "Advanced":
                    cell.fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")
                    cell.font = Font(bold=True, color="C62828")

    # Column widths
    widths = [8, 20, 15, 16, 28, 35, 32, 30, 45, 22]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.auto_filter.ref = f"A4:J{len(vocab_list)+4}"
    ws.freeze_panes = "B5"

    # ---------- Sheet 2: Summary ----------
    ws2 = wb.create_sheet("BCS & Job Prep Summary")
    ws2.merge_cells("A1:G1")
    ws2["A1"] = "DAILY STAR VOCABULARY ANALYSIS - BCS & JOB EXAM FOCUS"
    ws2["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws2["A1"].fill = navy
    ws2["A1"].alignment = Alignment(horizontal="center")

    # Level breakdown (static counts)
    level_counts = {"Basic": 0, "Intermediate": 0, "Advanced": 0}
    for item in vocab_list:
        lvl = item.get("level", "Basic")
        level_counts[lvl] = level_counts.get(lvl, 0) + 1

    ws2["A3"] = "VOCABULARY BREAKDOWN BY LEVEL"
    ws2["A3"].font = Font(bold=True, size=12)
    ws2["A4"] = "Difficulty Level"
    ws2["B4"] = "Word Count"
    ws2["C4"] = "% of Total"
    row = 5
    for lvl in ["Basic", "Intermediate", "Advanced"]:
        ws2[f"A{row}"] = lvl
        ws2[f"B{row}"] = level_counts.get(lvl, 0)
        ws2[f"C{row}"] = level_counts.get(lvl, 0) / len(vocab_list) if vocab_list else 0
        ws2[f"C{row}"].number_format = "0.0%"
        row += 1
    ws2[f"A{row}"] = "Total Vocabulary"
    ws2[f"B{row}"] = f"=SUM(B5:B7)"
    ws2[f"C{row}"] = "=SUM(C5:C7)"

    # Category breakdown (for display)
    cat_counts = {}
    for item in vocab_list:
        cat = item.get("category", "Uncategorized")
        cat_counts[cat] = cat_counts.get(cat, 0) + 1
    # Put on the right side (columns E-G)
    ws2["E3"] = "VOCABULARY DISTRIBUTION BY NEWSPAPER SECTOR"
    ws2["E3"].font = Font(bold=True, size=12)
    ws2["E4"] = "Sector / Category"
    ws2["F4"] = "Word Count"
    ws2["G4"] = "Target Job Exams"
    row = 5
    # Predefined order for clarity
    category_order = ["Politics & Governance", "Economy & Finance", "Law & Human Rights",
                      "Public Health & Env", "Crime & Legal Affairs", "Social & Policy Issues"]
    for cat in category_order:
        count = cat_counts.get(cat, 0)
        if count > 0:
            ws2[f"E{row}"] = cat
            ws2[f"F{row}"] = count
            # Target exams based on category (simplified)
            if "Politics" in cat:
                ws2[f"G{row}"] = "BCS Admin, Judiciary, Secretariat"
            elif "Economy" in cat or "Finance" in cat:
                ws2[f"G{row}"] = "Bangladesh Bank, Commercial Banks"
            elif "Law" in cat or "Rights" in cat:
                ws2[f"G{row}"] = "Judicial Service, Assistant Judge, Police"
            elif "Health" in cat or "Env" in cat:
                ws2[f"G{row}"] = "Medical Officer, BCS General Cadre"
            elif "Crime" in cat:
                ws2[f"G{row}"] = "BCS Police, ACC"
            else:
                ws2[f"G{row}"] = "BCS Written, General Knowledge"
            row += 1

    # Guidelines
    ws2.merge_cells("A10:G10")
    ws2["A10"] = "BANGLADESHI JOB EXAM PREPARATION GUIDELINES (BCS, BANK, NTRCA, PRIMARY)"
    ws2["A10"].font = Font(bold=True, size=12)
    guidelines = [
        "1. BCS Preliminary Strategy: Daily Star editorials contain high-frequency synonyms/antonyms asked directly in BCS Preliminary (20 Marks English). Focus on Advanced words.",
        "2. BCS Written Strategy: Words like 'Rupture', 'Reimagine', 'Transboundary', 'Exacerbate' enhance the standard of translation and passage writing in BCS Written.",
        "3. Bank Recruitment Exams: Bank exams heavily emphasize vocabulary in context, reading comprehension, and fill-in-the-blanks. Learn synonyms and precise nuances.",
        "4. Revision Technique: Practice making sentences using 5 new Daily Star vocabulary words daily to retain Bengali meaning and appropriate usage."
    ]
    for i, text in enumerate(guidelines, start=11):
        ws2.cell(row=i, column=1, value=text).alignment = Alignment(wrap_text=True)

    # ---------- Sheet 3: Practice Test ----------
    if mcqs and isinstance(mcqs, list) and all(isinstance(q, dict) for q in mcqs):
        ws3 = wb.create_sheet("BCS & Bank Practice Set")
        ws3.merge_cells("A1:G1")
        ws3["A1"] = "DAILY STAR VOCABULARY PRACTICE SET (JOB EXAM MODEL TEST)"
        ws3["A1"].font = Font(size=14, bold=True)
        ws3["A1"].alignment = Alignment(horizontal="center")

        ws3.merge_cells("A2:G2")
        ws3["A2"] = "Test your vocabulary knowledge for upcoming BCS Preliminary, Combined Bank Officers, and Primary Teacher exams."
        ws3["A2"].font = Font(size=10, italic=True)
        ws3["A2"].alignment = Alignment(horizontal="center")

        headers_mcq = ["Q. No", "Question / Contextual Sentence", "Option A", "Option B", "Option C", "Option D", "Correct Answer & Explanation"]
        for col, h in enumerate(headers_mcq, 1):
            cell = ws3.cell(row=4, column=col)
            cell.value = h
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")

        for i, q in enumerate(mcqs[:10], 5):
            ws3.cell(row=i, column=1, value=i-4).alignment = Alignment(horizontal="center")
            ws3.cell(row=i, column=2, value=safe_excel_value(q.get("question", ""))).alignment = Alignment(wrap_text=True)
            opts = q.get("options", [])
            for j, opt in enumerate(opts, 3):
                ws3.cell(row=i, column=j, value=safe_excel_value(opt))
            ws3.cell(row=i, column=7, value=safe_excel_value(q.get("answer", "") + " — " + q.get("explanation", ""))).alignment = Alignment(wrap_text=True)

        # Column widths
        ws3.column_dimensions["A"].width = 6
        ws3.column_dimensions["B"].width = 40
        ws3.column_dimensions["C"].width = 20
        ws3.column_dimensions["D"].width = 20
        ws3.column_dimensions["E"].width = 20
        ws3.column_dimensions["F"].width = 20
        ws3.column_dimensions["G"].width = 40
        ws3.freeze_panes = "A5"
    else:
        # If no MCQs, create a placeholder sheet
        ws3 = wb.create_sheet("BCS & Bank Practice Set")
        ws3.merge_cells("A1:F1")
        ws3["A1"] = "Practice Test – Not Generated (API issue)"
        ws3["A1"].font = Font(size=14, bold=True)
        ws3["A3"] = "No MCQs were generated due to a temporary issue. Try again tomorrow."

    wb.save(EXCEL_FILE)
    return EXCEL_FILE

# --------------------------------------------------------------
# BUILD PDF – MOBILE FRIENDLY
# --------------------------------------------------------------

def build_pdf(vocab_list, mcqs, summary, date_str):
    pdf_path = DATA_DIR / f"Vocabulary_{date_str}.pdf"

    # Try to use a Bengali font if available
    bengali_font = register_bengali_font()
    styles = getSampleStyleSheet()
    # Create a custom style for Bengali content
    bengali_style = ParagraphStyle(
        'BengaliStyle',
        parent=styles['Normal'],
        fontName=bengali_font,
        fontSize=10,
        leading=12
    )
    # Title style
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Title'],
        fontName=bengali_font,
        fontSize=16,
        alignment=1  # center
    )
    heading_style = ParagraphStyle(
        'HeadingStyle',
        parent=styles['Heading2'],
        fontName=bengali_font,
        fontSize=12,
        spaceAfter=6
    )

    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4,
                            rightMargin=72, leftMargin=72,
                            topMargin=72, bottomMargin=72)
    story = []

    # Title
    story.append(Paragraph("BCS & Bank Job Preparation – Daily Vocabulary", title_style))
    story.append(Paragraph(f"Date: {date_str}", bengali_style))
    story.append(Spacer(1, 0.2*inch))

    # Summary
    if summary:
        story.append(Paragraph("📰 Today's Summary", heading_style))
        # Replace newlines with <br/> for reportlab
        summary_clean = summary.replace('\n', '<br/>')
        story.append(Paragraph(summary_clean, bengali_style))
        story.append(Spacer(1, 0.2*inch))

    # Vocabulary list (first 50 items to keep PDF compact)
    story.append(Paragraph("📚 Vocabulary List (Top 50)", heading_style))
    data = [['Word', 'Bengali', 'POS', 'Example']]
    for item in vocab_list[:50]:
        data.append([
            item.get('word', ''),
            item.get('bengali', ''),
            item.get('pos', ''),
            item.get('example', '')[:60] + '...' if len(item.get('example', '')) > 60 else item.get('example', '')
        ])
    t = Table(data, colWidths=[1.2*inch, 1.5*inch, 0.8*inch, 2.5*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), bengali_font),
        ('FONTSIZE', (0,0), (-1,0), 10),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('BACKGROUND', (0,1), (-1,-1), colors.beige),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTNAME', (0,1), (-1,-1), bengali_font),
        ('FONTSIZE', (0,1), (-1,-1), 9),
    ]))
    story.append(t)
    story.append(PageBreak())

    # Practice Test
    if mcqs:
        story.append(Paragraph("📝 Practice Test (10 MCQs)", heading_style))
        for i, q in enumerate(mcqs[:10], 1):
            story.append(Paragraph(f"<b>{i}. {q['question']}</b>", bengali_style))
            for opt in q['options']:
                story.append(Paragraph(f"   {opt}", bengali_style))
            story.append(Paragraph(f"✅ <b>Answer:</b> {q['answer']} – {q['explanation']}", bengali_style))
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
