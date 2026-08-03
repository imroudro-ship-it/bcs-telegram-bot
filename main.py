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
    # Remove duplicates
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
- For each word provide: word, pos, level, bengali meaning, definition, synonyms, antonyms, example sentence, category.

### Task 2: Write a 5-7 bullet Bengali summary of the most important topics.

Return JSON with keys: "vocab_list" (array) and "bengali_summary" (string).
"""

    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": "You are an expert Bengali lexicographer. Output valid JSON."},
            {"role": "user", "content": prompt}
        ],
        model="llama-3.3-70b-versatile",
        temperature=0.3,
        response_format={"type": "json_object"},
    )
    data = json.loads(response.choices[0].message.content)
    return data.get("vocab_list", []), data.get("bengali_summary", "")

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def generate_mcqs(client, vocab_list):
    selected = random.sample(vocab_list, min(10, len(vocab_list)))
    word_list = [w["word"] for w in selected]

    prompt = f"""
Generate 10 MCQs from these words: {', '.join(word_list)}.
For each: question, options (A-D), answer (letter), explanation.
Return JSON array.
"""

    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": "You are an exam creator. Output valid JSON."},
            {"role": "user", "content": prompt}
        ],
        model="llama-3.3-70b-versatile",
        temperature=0.4,
        response_format={"type": "json_object"},
    )
    return json.loads(response.choices[0].message.content)

# ===================== BUILD EXCEL =====================
def build_excel(vocab_list, mcqs):
    wb = openpyxl.Workbook()

    # Sheet 1: Vocabulary
    ws = wb.active
    ws.title = "Vocabulary"

    # Styling
    navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    light_navy = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    even = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))

    # Title
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

    # Headers
    headers = ["SL", "Word", "POS", "Level", "Bengali", "Definition",
               "Synonyms", "Antonyms", "Example", "Category"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col)
        cell.value = h
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = navy
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Data
    for idx, item in enumerate(vocab_list, 5):
        ws.row_dimensions[idx].height = 26
        row_data = [
            item.get("sl", idx-4),
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
            cell = ws.cell(row=idx, column=col)
            cell.value = val
            cell.border = border
            cell.fill = even if idx % 2 == 0 else white
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if col in [1, 3, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Color by level
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

    # Auto-filter
    ws.auto_filter.ref = f"A4:J{len(vocab_list)+4}"
    ws.freeze_panes = "B5"

    # Sheet 2: Summary
    ws2 = wb.create_sheet("Summary")
    ws2.merge_cells("A1:G1")
    ws2["A1"] = "VOCABULARY BREAKDOWN"
    ws2["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws2["A1"].fill = navy
    ws2["A1"].alignment = Alignment(horizontal="center")

    # Level counts
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

    for i, q in enumerate(mcqs, 4):
        ws3.cell(row=i, column=1, value=i-3)
        ws3.cell(row=i, column=2, value=q["question"]).alignment = Alignment(wrap_text=True)
        for j, opt in enumerate(q["options"], 3):
            ws3.cell(row=i, column=j, value=opt)
        ws3.cell(row=i, column=7, value=q["answer"]).alignment = Alignment(horizontal="center")
        ws3.cell(row=i, column=8, value=q["explanation"]).alignment = Alignment(wrap_text=True)

    for col in range(1, 9):
        ws3.column_dimensions[get_column_letter(col)].width = 20 if col != 2 else 40
    ws3.freeze_panes = "A4"

    wb.save(EXCEL_FILE)
    return EXCEL_FILE

# ===================== CORE JOB LOGIC (used by both scheduled and interactive) =====================
async def run_daily_job(bot=None):
    """Generate vocabulary, summary, MCQs, build Excel, and send via bot (if provided)."""
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

    # If bot is provided, send the files and summary
    if bot:
        chat_id = TELEGRAM_CHAT_ID
        if chat_id:
            with open(EXCEL_FILE, "rb") as f:
                await bot.send_document(chat_id=chat_id, document=f, caption="📊 Daily Vocabulary Bank (auto-generated)")
            await bot.send_message(chat_id=chat_id, text=f"📰 **Summary**\n\n{summary}", parse_mode="Markdown")
        return "Sent successfully."
    else:
        # Just save and return, useful for local testing
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
        # Use the bot from context
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

# ===================== SCHEDULED JOB (called by APScheduler) =====================
async def scheduled_job(context: ContextTypes.DEFAULT_TYPE):
    await run_daily_job(bot=context.bot)

# ===================== MAIN =====================
def main():
    if not GROQ_API_KEY or not TELEGRAM_BOT_TOKEN:
        raise ValueError("Missing GROQ_API_KEY or TELEGRAM_BOT_TOKEN")

    # If running in GitHub Actions, run the scheduled job once and exit
    if os.environ.get("GITHUB_ACTIONS") == "true":
        # Create a bot instance to send messages
        from telegram import Bot
        bot = Bot(token=TELEGRAM_BOT_TOKEN)
        # Run the job asynchronously
        asyncio.run(run_daily_job(bot=bot))
        print("GitHub Actions job completed.")
        return

    # Otherwise, start the interactive bot
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
