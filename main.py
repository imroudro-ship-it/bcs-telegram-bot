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
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes
from tenacity import retry, stop_after_attempt, wait_exponential

# ===================== ENVIRONMENT VARIABLES =====================
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

HISTORY_FILE = "history.json"
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
SUMMARY_FILE = DATA_DIR / "summary.txt"
EXCEL_FILE = DATA_DIR / "Vocabulary_Bank.xlsx"

# ===================== CONFIG =====================
TIMEZONE = pytz.timezone("Asia/Dhaka")
RSS_FEEDS = [
    "https://www.thedailystar.net/rss.xml",
    "https://www.dhakatribune.com/feed",
    "https://www.tbsnews.net/rss.xml",
    "https://www.newagebd.net/rss.xml",
]

# ===================== HELPER: SAFELY CONVERT TO EXCEL STRING =====================
def safe_excel_value(value):
    if value is None:
        return ""
    if isinstance(value, list):
        return ", ".join(str(v) for v in value)
    if isinstance(value, dict):
        return str(value)
    return str(value)

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
    all_headlines = []
    for url in RSS_FEEDS:
        try:
            feed = feedparser.parse(url)
            for entry in feed.entries[:5]:
                all_headlines.append(entry.title)
        except:
            pass
    seen = set()
    unique = []
    for h in all_headlines:
        if h not in seen:
            seen.add(h)
            unique.append(h)
    return unique[:25]

# ===================== GENERATE VOCAB + SUMMARY =====================
def setup_groq():
    return groq.Groq(api_key=GROQ_API_KEY)

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def generate_vocab_and_summary(client, headlines, past_words):
    headlines_text = "\n".join([f"- {h}" for h in headlines])
    exclude = ", ".join(past_words[-50:])

    prompt = f"""
You are an expert BCS and Bank job exam mentor.

Today's headlines from Bangladeshi newspapers:
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

### Task 2: Write a 5-7 bullet Bengali summary of the most important topics.

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

    # **DEBUG: Print the first item to see keys**
    if vocab:
        print("📝 First vocab item keys:", vocab[0].keys())
        print("📝 First vocab item:", json.dumps(vocab[0], indent=2, ensure_ascii=False))

    # **Fallback: if keys are missing, try to map aliases**
    for item in vocab:
        # If 'bengali' is missing but 'bengali_meaning' exists, copy it
        if "bengali" not in item and "bengali_meaning" in item:
            item["bengali"] = item["bengali_meaning"]
        if "example" not in item and "example_sentence" in item:
            item["example"] = item["example_sentence"]
        # Ensure all keys exist with empty string as fallback
        for key in ["sl", "word", "pos", "level", "bengali", "definition", "synonyms", "antonyms", "example", "category"]:
            if key not in item:
                item[key] = ""

    return vocab, summary

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
        if len(validated) < 10:
            validated = []
        return validated
    except Exception as e:
        print(f"⚠️ MCQs parsing error: {e}")
        return []

# ===================== BUILD EXCEL (same as before, but uses safe_excel_value) =====================
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

# ===================== CORE JOB LOGIC =====================
async def run_daily_job(bot=None):
    client = setup_groq()
    headlines = fetch_headlines()
    if not headlines:
        return "No headlines found."

    past = load_history()
    vocab, summary = generate_vocab_and_summary(client, headlines, past)
    with open(SUMMARY_FILE, "w", encoding="utf-8") as f:
        f.write(summary)

    mcqs = generate_mcqs(client, vocab)

    build_excel(vocab, mcqs)
    save_history(past + [w["word"] for w in vocab])

    if bot:
        chat_id = TELEGRAM_CHAT_ID
        if chat_id:
            with open(EXCEL_FILE, "rb") as f:
                await bot.send_document(chat_id=chat_id, document=f, caption="📊 Daily Vocabulary Bank (auto-generated)")
            await bot.send_message(chat_id=chat_id, text=f"📰 **Summary**\n\n{summary}", parse_mode="Markdown")
        return "Sent successfully."
    else:
        return "Generated successfully."

# ===================== TELEGRAM COMMANDS =====================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🤖 **Daily Vocabulary Bot**\n\n"
        "/daily - Get today's vocabulary + summary\n"
        "/summary - Get Bengali summary\n"
        "/quiz - Get 10 MCQs\n\n"
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
        headlines = fetch_headlines()
        if not headlines:
            await update.message.reply_text("⚠️ No headlines.")
            return
        past = load_history()
        vocab, _ = generate_vocab_and_summary(client, headlines, past)
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
        asyncio.run(run_daily_job(bot=bot))
        print("GitHub Actions job completed.")
        return

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("daily", daily))
    app.add_handler(CommandHandler("summary", summary_cmd))
    app.add_handler(CommandHandler("quiz", quiz_cmd))

    scheduler = AsyncIOScheduler(timezone=TIMEZONE)
    scheduler.add_job(
        scheduled_job,
        trigger=CronTrigger(hour=15, minute=0, timezone=TIMEZONE),
        args=[app.context]
    )
    scheduler.start()

    print("🤖 Bot running. Commands: /start, /daily, /summary, /quiz")
    app.run_polling()

if __name__ == "__main__":
    main()
