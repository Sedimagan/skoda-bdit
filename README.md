# Gurudev Motors — Best Dealer in Town

Sales KPI score calculator for Škoda dealership "Gurudev Motors Pvt. Ltd."

This repo contains two versions of the same scoring logic:

- **`sales_score_calculator.py`** — the original Tkinter desktop app. Built into a
  Windows `.exe` automatically by `.github/workflows/build_windows.yml` on every push,
  and into the macOS `.app` locally via `build_windows.bat` / PyInstaller.
- **`webapp/`** — a Gradio web app with the same KPI/scoring/quarterly-average rules,
  backed by a shared Google Sheet so the whole team can enter and view scores from
  any browser. Deployed free on [Hugging Face Spaces](https://huggingface.co/spaces),
  auto-synced from this repo's `main` branch by
  `.github/workflows/deploy_hf_space.yml`.

## Running the web app locally

```bash
cd webapp
pip install -r requirements.txt
cp .env.example .env   # then fill in your own Sheet URL + service-account JSON
python app.py
```

See `webapp/.env.example` for what credentials are needed
(a Google service-account key + the target Sheet URL).
