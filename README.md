# Ai-ssistant 🧠🤖

Ai-ssistant is a lightweight Flask-based tool that helps you analyze batches of interview transcripts using an LLM.  
It extracts meaningful interviewee quotes, refines them, categorizes them, and exports everything into a clean, filterable Excel file.

---

## ✨ Features

- Batch upload of `.docx` interview transcripts
- Three-step LLM-powered analysis:
  1. Grammar & spelling correction
  2. Interviewee quote extraction (interviewer is ignored)
  3. Quote refinement and thematic categorization
- Automatic topic tagging:
  - Business Model
  - Market Outlook
  - Challenges
- Merged Excel output with formatting and filters
- Simple drag-and-drop web interface

---

## 🚀 Quick Start

### 1. Install dependencies

pip install flask werkzeug openpyxl python-docx

---

### 2. Project structure

Ai-ssistant/
├─ kitool.py  
├─ gpt4free/            # LLM client (or replace with your own)  
├─ inputs/              # input interview files (.docx)  
└─ outputs/             # generated Excel files  

---

### 3. Connect an LLM backend 🔌

Ai-ssistant requires a working LLM connection.

By default, it uses a `gpt4free` client:

from g4f.client import Client  
client = Client()

You may replace this with any LLM provider (OpenAI, Anthropic, etc.).  
If you do, update the LLM calls inside `process_entire_transcript()` in `kitool.py`.

⚠️ Without an LLM backend, the app will not produce outputs.

---

### 4. Run the application

python kitool.py

Then open your browser at:

http://127.0.0.1:5000

---

## 🗂️ Usage

### Web Interface

1. Drag & drop one or more `.docx` interview transcripts
2. Enter the interviewer name (their quotes will be ignored)
3. Click **Generate Excel**
4. Download the merged analysis file

---

### API Endpoint

POST /generate_excel

Form data:
- files: one or more `.docx` files
- interviewer: interviewer name

Response:
- Downloadable Excel file (`merged_output.xlsx`)

---

## 📁 File Locations

- Input files: uploads/
- Output files: outputs/merged_output.xlsx

You can change these folders via `UPLOAD_FOLDER` and `OUTPUT_FOLDER` in `kitool.py`.

---

## 📊 Excel Output Format

Columns:
- Interview File Name
- Timestamp
- Topic
- Quote

Topics:
- Business Model
- Market Outlook
- Challenges
- Uncategorized

---

## 🛠️ Customization

- Change accepted file types via `ALLOWED_EXTENSIONS`
- Adjust prompts and themes in `process_entire_transcript()`
- Swap LLM models or providers as needed

---

## 🐞 Troubleshooting

- No output? Verify your LLM backend is connected and responding
- Errors? Check console logs — intermediate LLM responses are printed for debugging

---

Happy interview analysis! 🚀
