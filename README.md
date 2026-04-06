# Docx Report Engine

A Python-based Word document report generator with a CLI script and a FastAPI REST API.

## Setup
pip install -r requirements.txt

## Run CLI
python3 main.py


Generates `output/final_report.docx`.

## Run API
python3 -m uvicorn app:app --reload --port 8001

Then POST to `http://127.0.0.1:8001/generate`:

{
  "title": "Quarterly Business Report",
  "quarter": "Q1 2026",
  "author": "Analytics Team",
  "header_size": 12,
  "footer_size": 10
}


## Project Structure

docx_report_engine/
├── config.py
├── main.py
├── app.py
├── content_generator.py
├── requirements.txt
└── modules/
    ├── styles.py
    ├── header_footer.py
    ├── formatting.py
    ├── sections.py
    └── borders.py
