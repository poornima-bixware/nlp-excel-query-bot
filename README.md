📈 NLP Excel Query Bot
A natural language interface for Excel files — ask questions in plain English, get instant answers, filters, and visualizations without writing a single formula or macro.

🧩 Problem
Excel files contain valuable business data, but extracting insights requires knowledge of formulas, pivot tables, or VBA. Non-technical users are locked out of their own data.

✅ Solution
An NLP-powered bot that understands Excel structure and lets users:

Ask questions in plain English
Filter, sort, and aggregate data conversationally
Generate charts from natural language descriptions
Export results to new Excel sheets automatically
🏗️ Architecture
User uploads Excel file (.xlsx)
        │
        ▼
  Excel Parser (openpyxl / pandas)
  Schema + header extraction
        │
        ▼
  NLP Layer (OpenAI GPT-4o)
  Query understanding + intent detection
        │
        ▼
  Pandas Query Engine
  (filter / groupby / sort / aggregate)
        │
        ├──► Text Answer
        │
        ├──► Chart (Matplotlib/Plotly)
        │
        └──► Excel Export (openpyxl)
🛠️ Tech Stack
Component	Technology
LLM	OpenAI GPT-4o
NLP Framework	LangChain
Excel Processing	openpyxl, pandas
Visualization	Matplotlib, Plotly
Backend	FastAPI
Language	Python 3.10+
📂 Project Structure
nlp-excel-query-bot/
├── parser/
│   ├── excel_loader.py         # Load and parse Excel files
│   └── schema_detector.py      # Detect headers, data types, sheet names
├── nlp/
│   ├── query_engine.py         # NLP → pandas query translation
│   └── prompts.py              # GPT-4o system prompts
├── output/
│   ├── chart_generator.py      # Visualization output
│   └── excel_exporter.py       # Export results to Excel
├── api/
│   └── main.py                 # FastAPI endpoints
├── sample_data/
│   └── sample_sales.xlsx       # Test dataset
├── requirements.txt
├── .env.example
└── README.md
⚙️ Setup
git clone https://github.com/poornimagithubrit/nlp-excel-query-bot
cd nlp-excel-query-bot
pip install -r requirements.txt
cp .env.example .env
# Add your OpenAI API key
🔑 Environment Variables
OPENAI_API_KEY=your_openai_api_key_here
🚀 Usage
uvicorn api.main:app --reload
Upload and Query:

# Upload Excel file
POST /upload    → returns file_id

# Ask a question
POST /query
{
  "file_id": "abc123",
  "question": "Show me total sales by region for Q1 2025"
}
Example Queries:

"Which salesperson had the highest revenue in January?"
"Show a pie chart of product category distribution"
"Filter all rows where profit margin is below 10%"
"What is the average order value by city?"
📊 Results
✅ Supports .xlsx and .csv files up to 100MB
✅ Multi-sheet Excel navigation
✅ Auto-detects column types (dates, currency, categories)
✅ Exports filtered results back to Excel
📦 Requirements
langchain
openai
fastapi
uvicorn
pandas
openpyxl
matplotlib
plotly
python-dotenv
python-multipart
📄 License
MIT License
