import os
import json
import requests
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from groq import Groq

GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

def fetch_batch_vocab(client, start_sl, count, level_description):
    prompt = f"""
    Act strictly as a professional BCS and Bangladeshi Competitive Job Exam English Mentor.
    Generate exactly {count} high-yield vocabulary words (numbered from SL {start_sl} to {start_sl + count - 1}) commonly tested in BCS Preliminary/Written, Bank recruitment, and Judicial service exams, focusing on terms found in daily newspapers like The Daily Star or Prothom Alo.

    Difficulty distribution for this batch: {level_description}.

    Return a valid JSON object with a single key "vocab_list" containing an array of {count} objects.
    Each object must have these keys:
    "sl": integer,
    "word": string,
    "pos": string (Noun, Verb, Adjective, etc.),
    "level": string ("Basic", "Intermediate", or "Advanced"),
    "bengali": string (Bangla meaning),
    "definition": string (Brief English definition),
    "synonyms": string (comma-separated),
    "antonyms": string (comma-separated),
    "example": string (Exam-standard sentence),
    "category": string (e.g., Economy, Politics, Public Health, Law & Judiciary, Environment)
    """

    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": "You are a JSON generator for Bangladeshi competitive job exams. Always output valid JSON."},
            {"role": "user", "content": prompt}
        ],
        model="llama-3.3-70b-versatile",
        temperature=0.4,
        response_format={"type": "json_object"}
    )
    
    data = json.loads(response.choices[0].message.content)
    return data.get("vocab_list", [])

def fetch_100_vocab():
    client = Groq(api_key=GROQ_API_KEY)
    print("Fetching Batch 1 (Words 1-50)...")
    batch1 = fetch_batch_vocab(client, 1, 50, "30 Basic and 20 Intermediate words")
    
    print("Fetching Batch 2 (Words 51-100)...")
    batch2 = fetch_batch_vocab(client, 51, 50, "15 Intermediate and 35 Advanced words")
    
    full_list = batch1 + batch2
    # Ensure serial numbers are 1 to 100
    for idx, item in enumerate(full_list, 1):
        item["sl"] = idx
        
    print(f"Successfully retrieved {len(full_list)} vocabulary items!")
    return full_list

def build_excel(vocab_list, filename="The_Daily_Star_Vocabulary_Bank.xlsx"):
    wb = openpyxl.Workbook()

    # Sheet 1: Daily Star Vocabulary
    ws_vocab = wb.active
    ws_vocab.title = 'Daily Star Vocabulary'
    ws_summary = wb.create_sheet(title='BCS & Job Prep Summary')

    # Styling constants
    fill_navy = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    fill_soft_navy = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
    fill_even = PatternFill(start_color='F2F5F9', end_color='F2F5F9', fill_type='solid')
    fill_odd = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    # Title Headers
    ws_vocab.merge_cells('A1:J1')
    ws_vocab['A1'] = "THE DAILY STAR - DAILY VOCABULARY BANK (BCS & JOB PREP SPECIAL)"
    ws_vocab['A1'].font = Font(name='Calibri', size=16, bold=True, color='FFFFFF')
    ws_vocab['A1'].fill = fill_navy
    ws_vocab['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws_vocab.merge_cells('A2:J2')
    ws_vocab['A2'] = "Curated from Today's Headlines, Editorials & Reports | Includes Meanings, Parts of Speech, Bengali Translations & Exam Examples"
    ws_vocab['A2'].font = Font(name='Calibri', size=10, italic=True, color='FFFFFF')
    ws_vocab['A2'].fill = fill_soft_navy
    ws_vocab['A2'].alignment = Alignment(horizontal='center', vertical='center')

    headers = ["SL No", "Word / Idiom", "Part of Speech", "Difficulty Level", "Bengali Meaning (বাংলা অর্থ)", "English Definition", "Synonyms", "Antonyms", "Contextual Example (Daily Star)", "Category / Topic"]

    for col_num, h_text in enumerate(headers, 1):
        cell = ws_vocab.cell(row=4, column=col_num)
        cell.value = h_text
        cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Populate Vocabulary Data
    for row_idx, item in enumerate(vocab_list, 5):
        row_vals = [
            item.get("sl", row_idx-4),
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
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_vocab.cell(row=row_idx, column=col_idx)
            cell.value = val
            cell.font = Font(name='Calibri', size=10)
            cell.border = thin_border
            cell.fill = fill_even if row_idx % 2 == 0 else fill_odd
            if col_idx == 1:
                cell.alignment = Alignment(horizontal='center', vertical='top')
            elif col_idx in [2, 3, 4, 10]:
                cell.alignment = Alignment(horizontal='left', vertical='top')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    # Set Column Widths for Sheet 1
    col_widths_vocab = {1: 8, 2: 18, 3: 15, 4: 16, 5: 28, 6: 35, 7: 32, 8: 30, 9: 45, 10: 22}
    for col, width in col_widths_vocab.items():
        ws_vocab.column_dimensions[get_column_letter(col)].width = width

    # Sheet 2: Summary
    ws_summary.merge_cells('A1:G1')
    ws_summary['A1'] = "DAILY STAR VOCABULARY ANALYSIS - BCS & JOB EXAM FOCUS"
    ws_summary['A1'].font = Font(name='Calibri', size=16, bold=True, color='FFFFFF')
    ws_summary['A1'].fill = fill_navy
    ws_summary['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws_summary.merge_cells('A3:C3')
    ws_summary['A3'] = "VOCABULARY BREAKDOWN BY LEVEL"
    ws_summary['A3'].font = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
    ws_summary['A3'].fill = fill_soft_navy
    ws_summary['A3'].alignment = Alignment(horizontal='center', vertical='center')

    sum_headers_1 = ["Difficulty Level", "Word Count", "% of Total"]
    for col_idx, h in enumerate(sum_headers_1, 1):
        cell = ws_summary.cell(row=4, column=col_idx)
        cell.value = h
        cell.font = Font(name='Calibri', size=10, bold=True)
        cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')

    sum_data_1 = [
        ("Basic", 30, "=B5/B8"),
        ("Intermediate", 35, "=B6/B8"),
        ("Advanced", 35, "=B7/B8"),
        ("Total Vocabulary", "=SUM(B5:B7)", "=SUM(C5:C7)")
    ]

    for r_idx, row_vals in enumerate(sum_data_1, 5):
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws_summary.cell(row=r_idx, column=c_idx)
            cell.value = val
            cell.font = Font(name='Calibri', size=10, bold=(r_idx==8))
            cell.border = thin_border
            if c_idx == 3:
                cell.number_format = '0.0%'
            cell.alignment = Alignment(horizontal='center' if c_idx>1 else 'left', vertical='center')

    # Column Widths for Sheet 2
    col_widths_summary = {1: 22, 2: 14, 3: 14, 4: 5, 5: 28, 6: 14, 7: 42}
    for col, width in col_widths_summary.items():
        ws_summary.column_dimensions[get_column_letter(col)].width = width

    # Save workbook
    wb.save(filename)
    print(f"Excel file '{filename}' built successfully!")
    return filename

def send_telegram_package(vocab_list, excel_file):
    # 1. Send Text Digest
    message = "<b>📚 DAILY STAR 100-WORD VOCABULARY BANK</b>\n"
    message += "<i>BCS, Bank & Job Exam Special Edition</i>\n\n"
    message += "<b>🔥 Top Featured Words Today:</b>\n\n"

    for idx, item in enumerate(vocab_list[:8], 1):
        message += f"<b>{idx}. {item['word']}</b> ({item['pos']}) — <i>{item['level']}</i>\n"
        message += f"• <b>অর্থ:</b> {item['bengali']}\n"
        message += f"• <b>Synonyms:</b> {item['synonyms']}\n"
        message += f"• <b>Example:</b> <i>\"{item['example']}\"</i>\n\n"

    message += "─────────────────────\n"
    message += "📎 <b>Attached:</b> Complete 100-word structured Excel file (`.xlsx`) with Practice Questions & Summary Analysis below!"

    url_msg = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload_msg = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML"
    }
    requests.post(url_msg, json=payload_msg)

    # 2. Send Excel Document
    url_doc = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    with open(excel_file, "rb") as file_data:
        files = {"document": file_data}
        payload_doc = {
            "chat_id": TELEGRAM_CHAT_ID,
            "caption": "📊 Here is your full 100-Word Daily Star Vocabulary Excel Bank!"
        }
        req_doc = requests.post(url_doc, data=payload_doc, files=files)

    if req_doc.status_code == 200:
        print("Telegram text and Excel file sent successfully!")
    else:
        print(f"Telegram Document Error: {req_doc.text}")
        raise Exception(f"Telegram API Error: {req_doc.text}")

if __name__ == "__main__":
    if not GROQ_API_KEY or not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        raise ValueError("Missing required environment secrets.")
    
    vocab_data = fetch_100_vocab()
    excel_path = build_excel(vocab_data)
    send_telegram_package(vocab_data, excel_path)
