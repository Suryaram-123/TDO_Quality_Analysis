# Trade Dimensions Operations — Store Matching Automation Engine (Streamlit)

## Setup & Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Project Structure
- `app.py` — Streamlit UI & file processing workflow (preserves exact FastAPI validation, temp handling, and UI)
- `engine.py` — Unchanged TDLINX red matching, R1–R6 orange rules, ownership mapping, and Summary sheet generator
- `static/style.css` — Exact UI theme and layout styles
- `.streamlit/config.toml` — Streamlit theme configuration
