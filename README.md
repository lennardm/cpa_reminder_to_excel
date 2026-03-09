# CPA Global PDF to Excel Converter

Converts CPA Global patent renewal reminder PDFs into formatted Excel spreadsheets.

## Requirements

Python 3.x with a virtual environment:

```bash
python3 -m venv venv
venv/bin/pip install pdfplumber openpyxl
```

## Usage

```bash
venv/bin/python convert.py <input.pdf> [output.xlsx]
```

If no output path is given, the Excel file is saved next to the PDF with the same name.

**Example:**

```bash
venv/bin/python convert.py Reminder_8248790.pdf
# → Reminder_8248790.xlsx
```

## Output

The generated Excel file contains one sheet with:

| Column | Description |
|---|---|
| Land | Country |
| Patent/Ans.nr. | Patent or application number |
| Innehavare | Patent holder |
| Er referens | Your reference code |
| År | Year |
| Förfallodag | Due date |
| Kostnad SEK | Cost in SEK |

A SUM formula for the total cost is added at the bottom of the Kostnad SEK column.
