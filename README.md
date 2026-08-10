# Factory Production System — Pro

Single-command cigarette manufacturing dashboard. Processes your Excel shift reports automatically — no manual data entry.

## Files in this folder

| File | Purpose |
|------|---------|
| `app.py` | Streamlit UI — run this |
| `data_processor.py` | Data engine: parsing, cleaning, metrics, exports |
| `security.py` | User auth (SHA-256, salted) |
| `requirements.txt` | Python dependencies |
| `README.md` | This file |

## One-time setup

1. Install Python from https://www.python.org/downloads/ (check "Add python.exe to PATH")
2. Open Command Prompt in this folder and run:

```
pip install -r requirements.txt
```

## Running it

```
streamlit run app.py
```

Login: default credentials are set on first run — check **Admin Panel → Change Your Own Password** immediately after setup.

## Key notes

- **Waste %** = (piece-count waste / output sticks) × 100. Cigarette wastage (grams) is tracked separately as its own KPI card — it is NOT included in the Waste % calculation because grams and sticks are different units.
- **OEE Performance** is relative to your best-performing machine in the selected date range — it is a directional ranking signal, not an audited OEE figure (true Performance requires a known rated machine speed, which your reports don't record).
- **Machine normalisation**: the stable machine name is used as the grouping key, not the equipment code, because codes vary between shifts for the same physical line in your source files.
- If `reportlab` isn't installed, PDF export is automatically disabled — everything else still works.
- To adjust the assumed shift length for OEE, edit `DEFAULT_SHIFT_MINUTES` near the top of `data_processor.py`.
