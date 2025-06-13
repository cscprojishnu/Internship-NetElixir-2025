# Internship-NetElixir-2025

# Google Ads Quality Assurance Automation Tool

This Django-based web application automates quality assurance (QA) checks on exported Google Ads data provided in Excel format. The tool validates best practices, campaign hygiene, and optimization signals based on predefined questions and maps each check to a specific sheet in the uploaded Excel file.

## 🔧 Features

- Upload Google Ads data as an Excel file (`.xls` or `.xlsx`)
- Runs 20+ automated QA checks (e.g., URL health, ad strength, keyword consistency, impression share loss)
- Each question is mapped to its own designated sheet
- Generates a downloadable Excel report with findings (including tables and charts)
- Displays clear error messages if:
  - Sheets are missing
  - File format is incorrect
  - Data is insufficient

## 🚀 How It Works

1. Upload a valid Excel file with relevant sheets (e.g., "Keyword Data", "Ad Data", "Conversions Tracking Data", etc.)
2. The app reads each question from `questions.txt`
3. For each question:
   - Finds the associated sheet using a static mapping
   - Runs a specific analysis function
   - Saves results and renders them on the web page and in the final Excel report
4. Displays download link to the generated report

## 📁 Expected Excel Sheet Names

Each question is tied to a specific sheet. The required sheet names include (but are not limited to):

- `Keyword Data`
- `Conversions Tracking Data`
- `Ad Data`
- `RSA Ad Data`
- `Campaign Data`
- `AdGroup Data`
- `Extensions Data`
- `Audiences`
- `DSA`
- `Campaigns`

Make sure your Excel file contains the above-named sheets as applicable.

## ✅ Requirements

- Python 3.8+
- Django 3.x or 4.x
- pandas
- openpyxl
- matplotlib
- requests
- beautifulsoup4

Install requirements using:

```bash
pip install -r requirements.txt
```

🧪 Running the App
```bash
python manage.py runserver
```
