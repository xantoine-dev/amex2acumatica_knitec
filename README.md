# AMEX Expense Claim Generator

A streamlined toolchain that converts American Express billing statements into per-cardholder claim files ready for Acumatica import. The system ships with a Streamlit web UI and a command-line interface, both powered by a shared pandas pipeline.

## Highlights
- **Modern pipeline**: shared logic in `amex_tool/pipeline.py` handles header detection, data cleaning, splitting by cardholder, and optional corporate card enrichment.
- **Dual interfaces**: use `streamlit_app.py` for an interactive experience or `Amex2acumatica_Refactored.py` for automated batch jobs.
- **Template-aware exports**: drop in a CSV/XLS/XLSX template to control output column order without editing code.
- **Legacy-free**: only the refactored entrypoints remain, so the codebase is portfolio-ready.

## Repository Layout
```
.
├── amex_tool/                # Shared processing pipeline
├── streamlit_app.py          # Streamlit UI built on the shared pipeline
├── Amex2acumatica_Refactored.py  # CLI entrypoint using the same pipeline
├── requirements.txt          # Runtime dependencies
├── Procfile / setup.sh       # Streamlit deployment helpers (Heroku style)
└── README.md                 # Project documentation
```

## Quick Start
1. **Install dependencies**
   ```bash
   python -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```
2. **Run the Streamlit app**
   ```bash
   streamlit run streamlit_app.py
   ```
   Upload an AMEX statement (`.csv`, `.xls`, `.xlsx`), optionally a template and corporate card map, then download the generated ZIP of claim files.

3. **CLI processing**
   ```bash
   python Amex2acumatica_Refactored.py \
     --statement path/to/amex_statement.csv \
     --output path/to/output_dir \
     --format excel \
     --corporate path/to/corporate_card_map.csv
   ```
   - `--format` options: `excel` (default) or `csv`.
   - `--template` can be added to control column order.

## Input Data
Sample datasets were removed from version control for confidentiality. Provide your own AMEX statement exports and optional corporate card mapping files when running the CLI or Streamlit app. If you need reusable fixtures, create synthetic data locally and keep it outside tracked Git history.

## Under The Hood
- **Header detection**: scans the first ~100 rows of AMEX statements to find the header row (no manual edits needed if AMEX changes spacing).
- **Corporate card matching**: merges by lower-cased last name extracted from the corporate card file.
- **Extensibility**: pipeline functions are split into load/clean/generate/save, making it simple to plug into other front-ends.

## Deployment Notes
- Streamlit-focused deployments can rely on `requirements.txt` plus the `Procfile` and `setup.sh` for Heroku-like environments.
- For local testing or CI, run the CLI against synthetic data you generate locally (and keep untracked) to validate outputs quickly.

## Future Ideas
- Add automated tests (pytest) with locally generated synthetic fixtures that remain untracked.
- Wire up a lint/test workflow (e.g., GitHub Actions).
- Extend the Streamlit UI with history, cloud storage, or additional validation rules.

---
Perfect for demonstrating data wrangling, automation, and web UI skills in a single portfolio project.
