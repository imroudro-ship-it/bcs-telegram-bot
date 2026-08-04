# --------------------------------------------------------------
# BUILD HTML – IBA COMPENDIUM STYLE (CATEGORIZED)
# --------------------------------------------------------------

def build_html(vocab_list, mcqs, summary, date_str):
    html_path = DATA_DIR / f"Vocabulary_{date_str}.html"

    # Group vocabulary by category
    categories = {}
    for item in vocab_list:
        cat = item.get('category', 'Miscellaneous')
        if cat not in categories:
            categories[cat] = []
        categories[cat].append(item)

    # Sort categories alphabetically
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
            mcq_html += f"""
            <div class="mcq">
                <p><strong>{i}. {q['question']}</strong></p>
                <ul>
                    {''.join(f'<li>{opt}</li>' for opt in q['options'])}
                </ul>
                <p><strong>✅ Answer:</strong> {q['answer']} – {q['explanation']}</p>
            </div>
            """
        mcq_html += "</div>"

    # Full HTML template
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
        {summary.replace('\n', '<br>')}
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
