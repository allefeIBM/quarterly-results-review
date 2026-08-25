# Quarterly Results Review — Presentation Generator

A local web application that reads KPI data from Excel spreadsheets and generates PowerPoint presentations for IBM Field Marketing quarterly reviews.

---

## Requirements

- **Python 3.9+** — [python.org/downloads](https://www.python.org/downloads/)
- **Tesseract OCR** *(optional — only needed for EPM screenshot parsing)*
  - macOS: `brew install tesseract`
  - Windows: [UB Mannheim installer](https://github.com/UB-Mannheim/tesseract/wiki)
  - Linux: `sudo apt install tesseract-ocr`

---

## Quick Start

### macOS (double-click)
Double-click the **`Open Quarterly Results Review.command`** file.  
The first run will install all Python dependencies automatically and open the app in your browser.

### Any platform (terminal)
```bash
# 1. Clone the repository
git clone https://github.com/YOUR_USERNAME/quarterly-results-review.git
cd quarterly-results-review

# 2. (Optional) create a virtual environment
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
python3 start.py
```

Then open **http://localhost:5050** in your browser.

---

## Configuration

Sensitive credentials (Airtable token, base ID, table ID) can be set via environment variables instead of hardcoding them:

```bash
export AIRTABLE_TOKEN="your_token_here"
export AIRTABLE_BASE_ID="your_base_id"
export AIRTABLE_TABLE_ID="your_table_id"
```

Copy `.env.example` to `.env` and fill in your values if you use a tool like [`python-dotenv`](https://pypi.org/project/python-dotenv/).

---

## Project Structure

```
.
├── app.py                                    # Flask backend — KPI loader + PPTX generator
├── start.py                                  # Startup helper (installs deps, opens browser)
├── index.html                                # Frontend UI
├── Open Quarterly Results Review.command     # macOS double-click launcher
├── requirements.txt                          # Python dependencies
├── Quarterly Results Review Template_8.18.pptx   # PPTX template (required)
├── *.xlsx                                    # KPI data spreadsheets (one per segment)
└── .env.example                              # Environment variable template
```

---

## Adding / Updating KPI Data

1. Place updated `.xlsx` files in the root of the repository (same folder as `app.py`).
2. Click **Reload Data** in the app UI (or visit `http://localhost:5050/api/reload`) to refresh without restarting the server.

---

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `GET`  | `/` | Serves the web UI |
| `GET`  | `/api/events` | Returns all events available in the spreadsheets |
| `GET`  | `/api/event/<brief>` | Returns KPI detail for a specific brief code |
| `POST` | `/api/generate` | Generates and downloads the PPTX for a brief (form fields: `brief`, optional `epm_image`) |
| `GET`  | `/api/reload` | Clears the in-memory cache and reloads all XLSX files |

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` |
| App doesn't open automatically | Visit `http://localhost:5050` manually |
| Tesseract not found | EPM screenshot parsing is disabled — install Tesseract to enable it |
| Port 5050 already in use | Change `PORT = 5050` in `start.py` and `port=5050` in `app.py` |

---

## License

Internal IBM tool — not for public distribution.
