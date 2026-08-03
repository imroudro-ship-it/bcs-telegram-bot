import html
import json
import os
import textwrap
import time
import feedparser
import openpyxl
import requests
from groq import Groq
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from tenacity import retry, stop_after_attempt, wait_exponential

# ===================== ENVIRONMENT VARIABLES =====================
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
HISTORY_FILE = "history.json"

# ===================== HELPER: LOAD/SAVE HISTORY (to avoid repeat words) =====================
def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r") as f:
            return json.load(f)
    return {"words": []}

def save_history(words_list):
    with open(HISTORY_FILE, "w") as f:
        json.dump({"words": words_list}, f)

# ===================== STEP 1: FETCH TODAY'S REAL NEWSPAPER HEADLINES (FREE RSS) =====================
def fetch_todays_headlines():
    print("📰 Fetching today's headlines from The Daily Star...")
    try:
        feed = feedparser.parse("https://www.thedailystar.net/rss.xml")
        headlines = []
        for entry in feed.entries[:15]:  # Get top 15 news
            headlines.append(entry.title)
        if headlines:
            print(f"✅ Found {len(headlines)} headlines today!")
            return headlines
        else:
            print("⚠️ RSS returned empty. Using fallback cache.")
            return None
    except Exception as e:
        print(f"⚠️ RSS fetch failed: {e}")
        return None

# ===================== STEP 2: FETCH VOCABULARY FROM GROQ (WITH RETRIES) =====================
@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def fetch_batch_vocab(client, start_sl, count, level_description, headlines, exclude_words):
    # Convert exclude list to a string for the prompt
    exclude_str = ", ".join(exclude_words[-50:])  # Only send last 50 to save tokens

    # Combine headlines into a clean list
    headlines_text = "\n".join([f"- {h}" for h in headlines])

    prompt = f"""
    Act strictly as a professional BCS and Bangladeshi Competitive Job Exam English Mentor.

    TODAY'S REAL NEWSPAPER HEADLINES (from The Daily Star):
    {headlines_text}

    Based SOLELY on the vocabulary and topics found in these TODAY'S headlines, generate exactly {count} high-yield vocabulary words (numbered from SL {start_sl} to {start_sl + count - 1}).

    Difficulty distribution for this batch: {level_description}.

    ⚠️ CRITICAL: DO NOT repeat these previously used words: {exclude_str}. Generate completely fresh words.

    CRITICAL INSTRUCTIONS FOR BENGALI TRANSLATION:
    1. Do NOT use literal, direct, or robotic word-for-word machine translation.
    2. Provide natural, idiomatic Bengali translations consistent with standard Bengali dictionaries.
    3. Ensure the Bengali meaning accurately reflects the specified Part of Speech (POS).

    Return a valid JSON object with a single key "vocab_list" containing an array of {count} objects.
    Each object must have these keys:
    "sl": integer,
    "word": string,
    "pos": string (Noun, Verb, Adjective, etc.),
    "level": string ("Basic", "Intermediate", or "Advanced"),
    "bengali": string (Natural Bangla meaning),
    "definition": string (Brief English definition),
    "synonyms": string (comma-separated),
    "antonyms": string (comma-separated),
    "example": string (Exam-standard sentence using the word),
    "category": string (e.g., Economy, Politics, Public Health, Law & Judiciary, Environment)
    """

    system_instruction = (
        "You are an expert Bengali lexicographer. Always output valid JSON. "
        "Never translate literally; use standard Bangla dictionary meanings."
    )

    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": system_instruction},
            {"role": "user", "content": prompt},
        ],
        model="llama-3.3-70b-versatile",
        temperature=0.3,
        response_format={"type": "json_object"},
    )

    data = json.loads(response.choices[0].message.content)
    return data.get("vocab_list", [])

# ===================== STEP 3: ORCHESTRATE THE 100 WORDS =====================
def fetch_100_vocab():
    # Load history to avoid repeats
    history = load_history()
    past_words = history.get("words", [])
    print(f"📚 Loaded {len(past_words)} historical words to avoid repetition.")

    # Fetch today's headlines
    headlines = fetch_todays_headlines()

    # ZOMBIE MODE: If no headlines, use cached history to send yesterday's data
    if not headlines:
        print("💀 ZOMBIE MODE ACTIVATED: Using yesterday's cached vocabulary.")
        if past_words:
            # We'll just rebuild the vocab list from history (but we don't have full data)
            # Fallback: generate generic words but tell user
            headlines = ["No internet connection today. Generating general exam vocabulary as fallback."]
        else:
            headlines = ["General BCS vocabulary preparation"]

    client = Groq(api_key=GROQ_API_KEY)
    print("⏳ Fetching Batch 1 (Words 1-50)...")
    batch1 = fetch_batch_vocab(
        client, 1, 50, "30 Basic and 20 Intermediate words", headlines, past_words
    )

    # Update past words with batch 1 to avoid crossing over into batch 2
    temp_words = [item.get("word", "") for item in batch1]
    combined_exclude = past_words + temp_words

    print("⏳ Fetching Batch 2 (Words 51-100)...")
    batch2 = fetch_batch_vocab(
        client, 51, 50, "15 Intermediate and 35 Advanced words", headlines, combined_exclude
    )

    full_list = batch1 + batch2
    for idx, item in enumerate(full_list, 1):
        item["sl"] = idx

    # Save today's words to history for tomorrow
    today_words = [item.get("word", "") for item in full_list]
    save_history(past_words + today_words)

    print(f"✅ Successfully retrieved {len(full_list)} vocabulary items!")
    return full_list

# ===================== STEP 4: BUILD EXCEL (YOUR ORIGINAL BEAUTIFUL CODE - KEPT INTACT) =====================
def build_excel(vocab_list, filename="The_Daily_Star_Vocabulary_Bank.xlsx"):
    wb = openpyxl.Workbook()

    # SHEET 1: Daily Star Vocabulary
    ws_vocab = wb.active
    ws_vocab.title = "Daily Star Vocabulary"
    ws_vocab.sheet_view.showGridLines = True

    fill_navy = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    fill_soft_navy = PatternFill(
        start_color="2F5597", end_color="2F5597", fill_type="solid"
    )
    fill_even = PatternFill(
        start_color="F2F5F9", end_color="F2F5F9", fill_type="solid"
    )
    fill_odd = PatternFill(
        start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
    )

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    diff_styles = {
        "Basic": {
            "fill": PatternFill(
                start_color="E8F5E9", end_color="E8F5E9", fill_type="solid"
            ),
            "font": Font(name="Calibri", size=10, bold=True, color="2E7D32"),
        },
        "Intermediate": {
            "fill": PatternFill(
                start_color="FFF8E1", end_color="FFF8E1", fill_type="solid"
            ),
            "font": Font(name="Calibri", size=10, bold=True, color="F57F17"),
        },
        "Advanced": {
            "fill": PatternFill(
                start_color="FFEBEE", end_color="FFEBEE", fill_type="solid"
            ),
            "font": Font(name="Calibri", size=10, bold=True, color="C62828"),
        },
    }

    # Title Headers
    ws_vocab.merge_cells("A1:J1")
    ws_vocab["A1"] = (
        "THE DAILY STAR - DAILY VOCABULARY BANK (BCS & JOB PREP SPECIAL)"
    )
    ws_vocab["A1"].font = Font(
        name="Calibri", size=16, bold=True, color="FFFFFF"
    )
    ws_vocab["A1"].fill = fill_navy
    ws_vocab["A1"].alignment = Alignment(
        horizontal="center", vertical="center"
    )

    ws_vocab.merge_cells("A2:J2")
    ws_vocab["A2"] = (
        "Curated from Today's Headlines, Editorials & Reports | Includes"
        " Meanings, Parts of Speech, Bengali Translations & Exam Examples"
    )
    ws_vocab["A2"].font = Font(
        name="Calibri", size=10, italic=True, color="FFFFFF"
    )
    ws_vocab["A2"].fill = fill_soft_navy
    ws_vocab["A2"].alignment = Alignment(
        horizontal="center", vertical="center"
    )

    headers = [
        "SL No",
        "Word / Idiom",
        "Part of Speech",
        "Difficulty Level",
        "Bengali Meaning (বাংলা অর্থ)",
        "English Definition",
        "Synonyms",
        "Antonyms",
        "Contextual Example (Daily Star)",
        "Category / Topic",
    ]

    ws_vocab.row_dimensions[4].height = 28
    for col_num, h_text in enumerate(headers, 1):
        cell = ws_vocab.cell(row=4, column=col_num)
        cell.value = h_text
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = fill_navy
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    # Populate Data
    for row_idx, item in enumerate(vocab_list, 5):
        ws_vocab.row_dimensions[row_idx].height = 26
        row_vals = [
            item.get("sl", row_idx - 4),
            item.get("word", ""),
            item.get("pos", ""),
            item.get("level", ""),
            item.get("bengali", ""),
            item.get("definition", ""),
            item.get("synonyms", ""),
            item.get("antonyms", ""),
            item.get("example", ""),
            item.get("category", ""),
        ]

        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_vocab.cell(row=row_idx, column=col_idx)
            cell.value = val
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border
            cell.fill = fill_even if row_idx % 2 == 0 else fill_odd

            if col_idx in [1, 3]:
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
            elif col_idx in [2, 10]:
                cell.alignment = Alignment(
                    horizontal="left", vertical="center"
                )
            else:
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )

            if col_idx == 4:
                level_str = str(val).strip().capitalize()
                if level_str in diff_styles:
                    cell.fill = diff_styles[level_str]["fill"]
                    cell.font = diff_styles[level_str]["font"]
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

    col_widths_vocab = {
        1: 8,
        2: 18,
        3: 15,
        4: 16,
        5: 28,
        6: 35,
        7: 32,
        8: 30,
        9: 45,
        10: 22,
    }
    for col, width in col_widths_vocab.items():
        ws_vocab.column_dimensions[get_column_letter(col)].width = width

    # SHEET 2: Summary
    ws_summary = wb.create_sheet(title="BCS & Job Prep Summary")
    ws_summary.sheet_view.showGridLines = True

    ws_summary.merge_cells("A1:G1")
    ws_summary["A1"] = "DAILY STAR VOCABULARY ANALYSIS - BCS & JOB EXAM FOCUS"
    ws_summary["A1"].font = Font(
        name="Calibri", size=16, bold=True, color="FFFFFF"
    )
    ws_summary["A1"].fill = fill_navy
    ws_summary["A1"].alignment = Alignment(
        horizontal="center", vertical="center"
    )

    ws_summary.merge_cells("A3:C3")
    ws_summary["A3"] = "VOCABULARY BREAKDOWN BY LEVEL"
    ws_summary["A3"].font = Font(
        name="Calibri", size=11, bold=True, color="FFFFFF"
    )
    ws_summary["A3"].fill = fill_soft_navy
    ws_summary["A3"].alignment = Alignment(
        horizontal="center", vertical="center"
    )

    sum_headers_1 = ["Difficulty Level", "Word Count", "% of Total"]
    for col_idx, h in enumerate(sum_headers_1, 1):
        cell = ws_summary.cell(row=4, column=col_idx)
        cell.value = h
        cell.font = Font(name="Calibri", size=10, bold=True)
        cell.fill = PatternFill(
            start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
        )
        cell.alignment = Alignment(horizontal="center", vertical="center")

    last_data_row = len(vocab_list) + 4
    sum_data_1 = [
        (
            "Basic",
            f"=COUNTIF('Daily Star Vocabulary'!D5:D{last_data_row}, \"Basic\")",
            "=B5/$B$8",
        ),
        (
            "Intermediate",
            (
                "=COUNTIF('Daily Star Vocabulary'!D5:D"
                f'{last_data_row}, "Intermediate")'
            ),
            "=B6/$B$8",
        ),
        (
            "Advanced",
            (
                '=COUNTIF(\'Daily Star Vocabulary\'!D5:D'
                f'{last_data_row}, "Advanced")'
            ),
            "=B7/$B$8",
        ),
        ("Total Vocabulary", "=SUM(B5:B7)", "=SUM(C5:C7)"),
    ]

    for r_idx, row_vals in enumerate(sum_data_1, 5):
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws_summary.cell(row=r_idx, column=c_idx)
            cell.value = val
            cell.font = Font(name="Calibri", size=10, bold=(r_idx == 8))
            cell.border = thin_border
            if c_idx == 3:
                cell.number_format = "0.0%"
            cell.alignment = Alignment(
                horizontal="center" if c_idx > 1 else "left", vertical="center"
            )

    ws_summary.merge_cells("E3:G3")
    ws_summary["E3"] = "SECTOR DISTRIBUTION & JOB EXAM RELEVANCE"
    ws_summary["E3"].font = Font(
        name="Calibri", size=11, bold=True, color="FFFFFF"
    )
    ws_summary["E3"].fill = fill_soft_navy
    ws_summary["E3"].alignment = Alignment(
        horizontal="center", vertical="center"
    )

    sum_headers_2 = [
        "Newspaper Sector",
        "Word Count",
        "Target Competitive Exams",
    ]
    for col_idx, h in enumerate(sum_headers_2, 5):
        cell = ws_summary.cell(row=4, column=col_idx)
        cell.value = h
        cell.font = Font(name="Calibri", size=10, bold=True)
        cell.fill = PatternFill(
            start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
        )
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sectors = [
        ("Economy & Finance", "Bangladesh Bank, Combined Banks, BCS Cadre"),
        ("Politics & Governance", "BCS Preliminary & Written, Ministry Jobs"),
        ("Law & Judiciary", "Judicial Service (BJS), ACC, Legal Officer"),
        ("Environment & Health", "Primary Teacher, NTRCA, Health Dept"),
    ]

    for idx, (sec_name, exam_target) in enumerate(sectors, 5):
        ws_summary.cell(
            row=idx, column=5, value=sec_name
        ).alignment = Alignment(horizontal="left", vertical="center")
        cell_cnt = ws_summary.cell(
            row=idx,
            column=6,
            value=(
                "=COUNTIF('Daily Star Vocabulary'!J5:J"
                f'{last_data_row}, "*{sec_name.split()[0]}*")'
            ),
        )
        cell_cnt.alignment = Alignment(horizontal="center", vertical="center")
        cell_target = ws_summary.cell(row=idx, column=7, value=exam_target)
        cell_target.alignment = Alignment(horizontal="left", vertical="center")

        for c_idx in range(5, 8):
            cell = ws_summary.cell(row=idx, column=c_idx)
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border

    col_widths_summary = {1: 22, 2: 14, 3: 14, 4: 4, 5: 25, 6: 14, 7: 45}
    for col, width in col_widths_summary.items():
        ws_summary.column_dimensions[get_column_letter(col)].width = width

    wb.save(filename)
    print(f"✅ Excel file '{filename}' built successfully!")
    return filename

# ===================== STEP 5: SEND TO TELEGRAM (WITH SAFE TEXT TRUNCATION) =====================
def send_telegram_package(vocab_list, excel_file):
    # 1. Send Text Digest (Truncated to avoid Telegram's 4096 character limit)
    message = "<b>📚 DAILY STAR 100-WORD VOCABULARY BANK</b>\n"
    message += "<i>BCS, Bank & Job Exam Special Edition</i>\n\n"
    message += "<b>🔥 Top Featured Words Today:</b>\n\n"

    for idx, item in enumerate(vocab_list[:8], 1):
        word = html.escape(str(item.get("word", "")))
        pos = html.escape(str(item.get("pos", "")))
        level = html.escape(str(item.get("level", "")))
        bengali = html.escape(str(item.get("bengali", "")))
        synonyms = html.escape(str(item.get("synonyms", "")))
        example = html.escape(str(item.get("example", "")))

        # Truncate example if too long (keeps message safe)
        example = textwrap.shorten(example, width=80, placeholder="...")

        message += f"<b>{idx}. {word}</b> ({pos}) — <i>{level}</i>\n"
        message += f"• <b>অর্থ:</b> {bengali}\n"
        message += f"• <b>Synonyms:</b> {synonyms}\n"
        message += f'• <b>Example:</b> <i>"{example}"</i>\n\n'

    message += "─────────────────────\n"
    message += (
        "📎 <b>Attached:</b> Complete 100-word structured Excel file (`.xlsx`)"
        " with Summary Analysis below!"
    )

    url_msg = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload_msg = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML",
    }

    try:
        res_msg = requests.post(url_msg, json=payload_msg, timeout=30)
        print("✅ Telegram Text sent:", res_msg.status_code)
    except Exception as e:
        print(f"⚠️ Failed to send text: {e}")

    # 2. Send Excel Document
    url_doc = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    with open(excel_file, "rb") as file_data:
        files = {"document": file_data}
        payload_doc = {
            "chat_id": TELEGRAM_CHAT_ID,
            "caption": (
                "📊 Here is your full 100-Word Daily Star Vocabulary Excel Bank!"
            ),
        }
        try:
            req_doc = requests.post(url_doc, data=payload_doc, files=files, timeout=60)
            print("✅ Telegram File sent:", req_doc.status_code)
        except Exception as e:
            print(f"⚠️ Failed to send file: {e}")

# ===================== ENTRY POINT =====================
if __name__ == "__main__":
    if not GROQ_API_KEY or not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        raise ValueError(
            "❌ Missing required environment secrets (GROQ_API_KEY,"
            " TELEGRAM_BOT_TOKEN, or TELEGRAM_CHAT_ID)."
        )

    print("🚀 Starting Daily Newspaper Vocabulary Automation...")
    try:
        vocab_data = fetch_100_vocab()
        excel_path = build_excel(vocab_data)
        send_telegram_package(vocab_data, excel_path)
        print("🎉 Automation completed successfully!")
    except Exception as e:
        print(f"🔥 Critical error: {e}")
        # Even if it fails, try to send a fallback error message to Telegram
        try:
            url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
            requests.post(url, json={
                "chat_id": TELEGRAM_CHAT_ID,
                "text": f"⚠️ Daily Vocab Bot crashed. Error: {e}",
                "parse_mode": "HTML"
            })
        except:
            pass
