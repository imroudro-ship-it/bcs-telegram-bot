#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import asyncio
import json
import os
import random
from datetime import datetime, timedelta
from pathlib import Path

import feedparser
import groq
import openpyxl
import pytz
import requests
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes
from tenacity import retry, stop_after_attempt, wait_exponential
import jinja2

# ===================== ENVIRONMENT VARIABLES =====================
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

HISTORY_FILE = "history.json"
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
SUMMARY_FILE = DATA_DIR / "summary.txt"
EXCEL_FILE = DATA_DIR / "Vocabulary_Bank.xlsx"
DATABASE_FILE = DATA_DIR / "database.json"
DASHBOARD_FILE = DATA_DIR / "dashboard.html"

# ===================== CONFIG =====================
TIMEZONE = pytz.timezone("Asia/Dhaka")
RSS_FEEDS = {
    "The Daily Star": "https://www.thedailystar.net/rss.xml",
    "Dhaka Tribune": "https://www.dhakatribune.com/feed",
    "The Business Standard": "https://www.tbsnews.net/rss.xml",
    "New Age": "https://www.newagebd.net/rss.xml",
}

# ===================== HELPER: SAFELY CONVERT TO EXCEL STRING =====================
def safe_excel_value(value):
    if value is None:
        return ""
    if isinstance(value, list):
        return ", ".join(str(v) for v in value)
    if isinstance(value, dict):
        return str(value)
    return str(value)

# ===================== DATABASE =====================
def load_database():
    if DATABASE_FILE.exists():
        with open(DATABASE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"entries": []}

def save_database(db):
    with open(DATABASE_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=2, ensure_ascii=False)

# ===================== HISTORY =====================
def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r") as f:
            return json.load(f).get("words", [])
    return []

def save_history(words):
    with open(HISTORY_FILE, "w") as f:
        json.dump({"words": words}, f)

# ===================== FETCH HEADLINES =====================
def fetch_headlines():
    all_entries = []
    for source, url in RSS_FEEDS.items():
        try:
            feed = feedparser.parse(url)
            for entry in feed.entries[:5]:
                all_entries.append({
                    "headline": entry.title,
                    "source": source
                })
        except:
            pass
    seen = set()
    unique = []
    for entry in all_entries:
        if entry["headline"] not in seen:
            seen.add(entry["headline"])
            unique.append(entry)
    return unique[:25]

# ===================== GENERATE VOCAB + SUMMARY =====================
def setup_groq():
    return groq.Groq(api_key=GROQ_API_KEY)

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
  - "sl" (number)
  - "word" (string)
  - "pos" (string, e.g., Noun, Verb, Adjective)
  - "level" (string, one of: Basic, Intermediate, Advanced)
  - "bengali" (string, natural Bengali meaning)
  - "definition" (string, brief English definition)
  - "synonyms" (comma-separated string)
  - "antonyms" (comma-separated string)
  - "example" (string, exam-standard sentence using the word)
  - "category" (string, e.g., Economy, Politics, Law, Environment, Health)

### Task 2: Write a 5-7 bullet Bengali summary of the most important topics. Include the newspaper names in the summary.

Return JSON with keys: "vocab_list" (array) and "bengali_summary" (string).
"""

    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": "You are an expert Bengali lexicographer. Always output valid JSON. Use exactly the keys specified."},
            {"role": "user", "content": prompt}
        ],
        model="llama-3.3-70b-versatile",
        temperature=0.3,
        response_format={"type": "json_object"},
    )
    data = json.loads(response.choices[0].message.content)
    vocab = data.get("vocab_list", [])
    summary = data.get("bengali_summary", "")

    for item in vocab:
        if "bengali" not in item and "bengali_meaning" in item:
            item["bengali"] = item["bengali_meaning"]
        if "example" not in item and "example_sentence" in item:
            item["example"] = item["example_sentence"]
        for key in ["sl", "word", "pos", "level", "bengali", "definition", "synonyms", "antonyms", "example", "category"]:
            if key not in item:
                item[key] = ""

    if not summary or len(summary.strip()) < 10:
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

# ===================== GENERATE MCQs =====================
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

# ===================== BUILD EXCEL =====================
def build_excel(vocab_list, mcqs):
    wb = openpyxl.Workbook()

    # Sheet 1: Vocabulary
    ws = wb.active
    ws.title = "Vocabulary"

    navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    light_navy = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    even = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))

    ws.merge_cells("A1:J1")
    ws["A1"] = "DAILY VOCABULARY BANK - BCS & JOB PREP"
    ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = navy
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:J2")
    ws["A2"] = datetime.now(TIMEZONE).strftime("%A, %d %B %Y")
    ws["A2"].font = Font(size=10, italic=True, color="FFFFFF")
    ws["A2"].fill = light_navy
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["SL", "Word", "POS", "Level", "Bengali", "Definition",
               "Synonyms", "Antonyms", "Example", "Category"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col)
        cell.value = h
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = navy
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for idx, item in enumerate(vocab_list, 5):
        ws.row_dimensions[idx].height = 26
        row_data = [
            item.get("sl", idx-4),
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

    widths = [8, 20, 15, 16, 28, 35, 32, 30, 45, 22]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.auto_filter.ref = f"A4:J{len(vocab_list)+4}"
    ws.freeze_panes = "B5"

    # Sheet 2: Summary
    ws2 = wb.create_sheet("Summary")
    ws2.merge_cells("A1:G1")
    ws2["A1"] = "VOCABULARY BREAKDOWN"
    ws2["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws2["A1"].fill = navy
    ws2["A1"].alignment = Alignment(horizontal="center")

    ws2["A3"] = "Level"
    ws2["B3"] = "Count"
    ws2["C3"] = "%"
    for r, lvl in enumerate(["Basic", "Intermediate", "Advanced"], 4):
        ws2[f"A{r}"] = lvl
        ws2[f"B{r}"] = f"=COUNTIF(Vocabulary!D:D,\"{lvl}\")"
        ws2[f"C{r}"] = f"=B{r}/$B$7"
        ws2[f"C{r}"].number_format = "0.0%"
    ws2["A7"] = "Total"
    ws2["B7"] = "=SUM(B4:B6)"
    ws2["C7"] = "=SUM(C4:C6)"

    # Sheet 3: Practice Test
    if mcqs and isinstance(mcqs, list) and all(isinstance(q, dict) for q in mcqs):
        ws3 = wb.create_sheet("Practice Test")
        ws3.merge_cells("A1:F1")
        ws3["A1"] = "10 MCQs - Test Yourself"
        ws3["A1"].font = Font(size=14, bold=True)
        ws3["A1"].alignment = Alignment(horizontal="center")

        mcq_headers = ["#", "Question", "A", "B", "C", "D", "Answer", "Explanation"]
        for col, h in enumerate(mcq_headers, 1):
            cell = ws3.cell(row=3, column=col)
            cell.value = h
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        for i, q in enumerate(mcqs[:10], 4):
            ws3.cell(row=i, column=1, value=i-3)
            ws3.cell(row=i, column=2, value=safe_excel_value(q.get("question", ""))).alignment = Alignment(wrap_text=True)
            opts = q.get("options", [])
            for j, opt in enumerate(opts, 3):
                ws3.cell(row=i, column=j, value=safe_excel_value(opt))
            ws3.cell(row=i, column=7, value=safe_excel_value(q.get("answer", ""))).alignment = Alignment(horizontal="center")
            ws3.cell(row=i, column=8, value=safe_excel_value(q.get("explanation", ""))).alignment = Alignment(wrap_text=True)

        for col in range(1, 9):
            ws3.column_dimensions[get_column_letter(col)].width = 20 if col != 2 else 40
        ws3.freeze_panes = "A4"
    else:
        ws3 = wb.create_sheet("Practice Test")
        ws3.merge_cells("A1:F1")
        ws3["A1"] = "Practice Test - Not Generated (API issue)"
        ws3["A1"].font = Font(size=14, bold=True)
        ws3["A1"].alignment = Alignment(horizontal="center")
        ws3["A3"] = "No MCQs were generated due to a temporary issue. Try again tomorrow."

    wb.save(EXCEL_FILE)
    return EXCEL_FILE

# ===================== BUILD WEEKLY EXCEL =====================
def build_weekly_excel(unique_words, entries):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Weekly Vocabulary"

    # Headers
    headers = ["SL", "Word", "POS", "Level", "Bengali", "Definition", "Synonyms", "Antonyms", "Example", "Category"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col)
        cell.value = h
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    for idx, item in enumerate(unique_words, 2):
        row_data = [
            idx-1,
            item.get("word", ""),
            item.get("pos", ""),
            item.get("level", ""),
            item.get("bengali", ""),
            item.get("definition", ""),
            item.get("synonyms", ""),
            item.get("antonyms", ""),
            item.get("example", ""),
            item.get("category", "")
        ]
        for col, val in enumerate(row_data, 1):
            ws.cell(row=idx, column=col, value=safe_excel_value(val))

    # Summary sheet
    ws2 = wb.create_sheet("Weekly Stats")
    ws2["A1"] = "Week Summary"
    ws2["A1"].font = Font(bold=True, size=14)
    ws2["A3"] = "Date Range:"
    ws2["B3"] = f"{entries[0]['date'][:10]} to {entries[-1]['date'][:10]}"
    ws2["A4"] = "Total Unique Words:"
    ws2["B4"] = len(unique_words)

    weekly_file = DATA_DIR / f"weekly_report_{datetime.now(TIMEZONE).strftime('%Y-%m-%d')}.xlsx"
    wb.save(weekly_file)
    return weekly_file

# ===================== GENERATE DASHBOARD =====================
def generate_dashboard(db):
    entries = db.get("entries", [])
    if not entries:
        return

    # Prepare data for charts
    dates = [e["date"][:10] for e in entries[-30:]]  # last 30 days
    counts = [len(e.get("words", [])) for e in entries[-30:]]

    # Level distribution (last 7 days)
    level_counts = {"Basic": 0, "Intermediate": 0, "Advanced": 0}
    for e in entries[-7:]:
        for w in e.get("words", []):
            lvl = w.get("level", "Basic")
            level_counts[lvl] = level_counts.get(lvl, 0) + 1

    # Category counts (last 7 days)
    cat_counts = {}
    for e in entries[-7:]:
        for w in e.get("words", []):
            cat = w.get("category", "Uncategorized")
            cat_counts[cat] = cat_counts.get(cat, 0) + 1

    # Build HTML using jinja2
    template = """<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Vocabulary Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body { font-family: Arial, sans-serif; background: #f4f6f9; margin: 20px; }
        .container { max-width: 1200px; margin: auto; background: white; padding: 20px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #1F4E78; }
        .chart-row { display: flex; flex-wrap: wrap; gap: 20px; margin-top: 20px; }
        .chart-box { flex: 1; min-width: 300px; background: #fff; padding: 15px; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,0.05); }
        canvas { max-height: 250px; width: 100% !important; }
        .stats { display: flex; gap: 30px; flex-wrap: wrap; margin-bottom: 20px; }
        .stat-card { background: #1F4E78; color: white; padding: 15px 25px; border-radius: 8px; }
        .stat-card span { font-size: 28px; font-weight: bold; display: block; }
    </style>
</head>
<body>
<div class="container">
    <h1>📊 Vocabulary Progress Dashboard</h1>
    <div class="stats">
        <div class="stat-card">Total Words Learned <span>{{ total_words }}</span></div>
        <div class="stat-card">Days Active <span>{{ days_active }}</span></div>
        <div class="stat-card">Words This Week <span>{{ week_words }}</span></div>
    </div>
    <div class="chart-row">
        <div class="chart-box"><h3>Daily Words (last 30 days)</h3><canvas id="dailyChart"></canvas></div>
        <div class="chart-box"><h3>Level Distribution (last 7 days)</h3><canvas id="levelChart"></canvas></div>
    </div>
    <div class="chart-row">
        <div class="chart-box"><h3>Top Categories (last 7 days)</h3><canvas id="categoryChart"></canvas></div>
    </div>
</div>
<script>
    const dailyData = {{ daily_data|tojson }};
    const levelData = {{ level_data|tojson }};
    const categoryData = {{ category_data|tojson }};

    new Chart(document.getElementById('dailyChart'), {
        type: 'bar',
        data: {
            labels: dailyData.labels,
            datasets: [{ label: 'Words per day', data: dailyData.values, backgroundColor: '#1F4E78' }]
        },
        options: { responsive: true, plugins: { legend: { display: false } } }
    });

    new Chart(document.getElementById('levelChart'), {
        type: 'pie',
        data: {
            labels: levelData.labels,
            datasets: [{ data: levelData.values, backgroundColor: ['#2E7D32', '#F57F17', '#C62828'] }]
        }
    });

    new Chart(document.getElementById('categoryChart'), {
        type: 'bar',
        data: {
            labels: categoryData.labels,
            datasets: [{ label: 'Words', data: categoryData.values, backgroundColor: '#2F5597' }]
        }
    });
</script>
</body>
</html>
"""

    total_words = len(set(w["word"] for e in entries for w in e.get("words", [])))
    days_active = len(entries)
    week_words = sum(len(e.get("words", [])) for e in entries[-7:])

    env = jinja2.Environment()
    html = env.from_string(template).render(
        total_words=total_words,
        days_active=days_active,
        week_words=week_words,
        daily_data={"labels": dates, "values": counts},
        level_data={"labels": list(level_counts.keys()), "values": list(level_counts.values())},
        category_data={"labels": list(cat_counts.keys())[:10], "values": list(cat_counts.values())[:10]}
    )

    with open(DASHBOARD_FILE, "w", encoding="utf-8") as f:
        f.write(html)

# ===================== CORE JOB LOGIC =====================
@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
async def run_daily_job(bot=None):
    client = setup_groq()
    headlines_data = fetch_headlines()
    if not headlines_data:
        return "No headlines found."

    past = load_history()
    vocab, summary = generate_vocab_and_summary(client, headlines_data, past)
    with open(SUMMARY_FILE, "w", encoding="utf-8") as f:
        f.write(summary)

    mcqs = generate_mcqs(client, vocab)
    build_excel(vocab, mcqs)
    save_history(past + [w["word"] for w in vocab])

    # Save to database
    db = load_database()
    entry = {
        "date": datetime.now(TIMEZONE).isoformat(),
        "word_count": len(vocab),
        "levels": {"Basic": 0, "Intermediate": 0, "Advanced": 0},
        "categories": {},
        "words": vocab,
        "summary": summary
    }
    for w in vocab:
        lvl = w.get("level", "Basic")
        entry["levels"][lvl] = entry["levels"].get(lvl, 0) + 1
        cat = w.get("category", "Uncategorized")
        entry["categories"][cat] = entry["categories"].get(cat, 0) + 1
    db["entries"].append(entry)
    save_database(db)

    # Generate dashboard
    generate_dashboard(db)

    # Weekly summary on Sundays
    if datetime.now(TIMEZONE).weekday() == 6:  # Sunday
        last_7 = db["entries"][-7:]
        all_words = []
        for e in last_7:
            all_words.extend(e.get("words", []))
        seen = set()
        unique_words = []
        for w in all_words:
            if w["word"] not in seen:
                seen.add(w["word"])
                unique_words.append(w)
        if unique_words:
            weekly_file = build_weekly_excel(unique_words, last_7)
            if bot:
                with open(weekly_file, "rb") as f:
                    await bot.send_document(chat_id=TELEGRAM_CHAT_ID, document=f, caption="📊 Weekly Vocabulary Review")

    if bot:
        chat_id = TELEGRAM_CHAT_ID
        if chat_id:
            with open(EXCEL_FILE, "rb") as f:
                await bot.send_document(chat_id=chat_id, document=f, caption="📊 Daily Vocabulary Bank (auto-generated)")
            await bot.send_message(chat_id=chat_id, text=f"📰 **Today's Summary**\n\n{summary}", parse_mode="Markdown")
        return "Sent successfully."
    else:
        return "Generated successfully."

# ===================== TELEGRAM COMMANDS =====================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🤖 **Daily Vocabulary Bot**\n\n"
        "/daily - Get today's vocabulary + summary\n"
        "/summary - Get Bengali summary\n"
        "/quiz - Get 10 MCQs\n"
        "/last 3 - Show last 3 days' words\n\n"
        "Auto-sends daily at 3 PM Asia/Dhaka.",
        parse_mode="Markdown"
    )

async def daily(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = await update.message.reply_text("⏳ Generating...")
    try:
        result = await run_daily_job(bot=context.bot)
        await update.message.reply_text("✅ " + result)
        await msg.delete()
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")

async def summary_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if SUMMARY_FILE.exists():
        with open(SUMMARY_FILE, "r", encoding="utf-8") as f:
            await update.message.reply_text(f"📰 **Summary**\n\n{f.read()}", parse_mode="Markdown")
    else:
        await update.message.reply_text("No summary yet. Use /daily first.")

async def quiz_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("⏳ Generating quiz...")
    try:
        client = setup_groq()
        headlines_data = fetch_headlines()
        if not headlines_data:
            await update.message.reply_text("⚠️ No headlines.")
            return
        past = load_history()
        vocab, _ = generate_vocab_and_summary(client, headlines_data, past)
        mcqs = generate_mcqs(client, vocab)
        if not mcqs:
            await update.message.reply_text("⚠️ Could not generate MCQs. Please try again later.")
            return
        reply = "📝 **Practice Test**\n\n"
        for i, q in enumerate(mcqs, 1):
            reply += f"**{i}. {q['question']}**\n"
            for opt in q["options"]:
                reply += f"   {opt}\n"
            reply += f"✅ **Answer:** {q['answer']}\n"
            reply += f"💡 {q['explanation']}\n\n"
        await update.message.reply_text(reply, parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"❌ Quiz error: {e}")

async def last_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    n = 3
    if args and args[0].isdigit():
        n = min(int(args[0]), 7)
    db = load_database()
    entries = db["entries"][-n:]
    if not entries:
        await update.message.reply_text("No historical data yet. Use /daily first.")
        return
    reply = f"📜 **Last {len(entries)} days**\n\n"
    for e in reversed(entries):
        date = e["date"][:10]
        words = e.get("words", [])
        top5 = [w["word"] for w in words[:5]]
        reply += f"**{date}** – {len(words)} words\n"
        reply += f"   Top: {', '.join(top5)}\n"
        if e.get("summary"):
            summary_short = e["summary"][:100] + "..." if len(e["summary"]) > 100 else e["summary"]
            reply += f"   📰 {summary_short}\n"
        reply += "\n"
    await update.message.reply_text(reply, parse_mode="Markdown")

# ===================== SCHEDULED JOB =====================
async def scheduled_job(context: ContextTypes.DEFAULT_TYPE):
    await run_daily_job(bot=context.bot)

# ===================== MAIN =====================
def main():
    if not GROQ_API_KEY or not TELEGRAM_BOT_TOKEN:
        raise ValueError("Missing GROQ_API_KEY or TELEGRAM_BOT_TOKEN")

    if os.environ.get("GITHUB_ACTIONS") == "true":
        from telegram import Bot
        bot = Bot(token=TELEGRAM_BOT_TOKEN)
        try:
            asyncio.run(run_daily_job(bot=bot))
            print("GitHub Actions job completed.")
        except Exception as e:
            error_msg = f"❌ Daily job failed after 3 attempts: {str(e)[:200]}"
            bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=error_msg)
            print(error_msg)
        return

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("daily", daily))
    app.add_handler(CommandHandler("summary", summary_cmd))
    app.add_handler(CommandHandler("quiz", quiz_cmd))
    app.add_handler(CommandHandler("last", last_cmd))

    scheduler = AsyncIOScheduler(timezone=TIMEZONE)
    scheduler.add_job(
        scheduled_job,
        trigger=CronTrigger(hour=15, minute=0, timezone=TIMEZONE),
        args=[app.context]
    )
    scheduler.start()

    print("🤖 Bot running. Commands: /start, /daily, /summary, /quiz, /last")
    app.run_polling()

if __name__ == "__main__":
    main()
