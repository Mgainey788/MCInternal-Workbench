# MedComms Internal Workbench

A Streamlit app for MedComms source attribution, copyright QA, and reference fact-checking.

### How to run it on your own machine

1. Install the requirements

   ```
   $ pip install -r requirements.txt
   ```

2. Run the app

   ```
   $ streamlit run streamlit_app.py
   ```

### Requirements

| Package | Purpose |
|---|---|
| `streamlit` | Web app framework |
| `pandas` | Data manipulation and Excel/CSV export |
| `requests` | HTTP requests for URL-based source fetching |
| `beautifulsoup4` | HTML parsing |
| `lxml` | XML/HTML parser backend |
| `pypdf` | PDF text extraction |
| `python-pptx` | PowerPoint file parsing |
| `openpyxl` | Excel file read/write |
| `sentence-transformers` | Semantic embedding model (SPECTER2 / MiniLM) |
| `qdrant-client` | In-memory vector store for semantic search |
