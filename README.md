# Xelix Brand Aligner — web app

Upload one Xelix deck, align it to the brand guidelines, download the result.

## What it does (one pass, per slide)
- **Embeds Barlow** (Regular/Bold/Black) in the .pptx so it renders correctly on
  any machine — this is what stops headers, stat numbers and cards from
  overflowing when the viewer doesn't have the font installed.
- Sets **Barlow** typography (titles in Barlow Black).
- Ensures **one** wordmark bottom-right, coloured for the background behind that
  corner (white on dark, blue on light). Existing corner logos are kept, not doubled.
- **Repairs contrast locally**: white-on-dark and dark-on-light text inside cells
  and panels is fixed without touching correctly-coloured or mid-tone brand text.
- Backgrounds and layouts are left intact.

## Run locally
    pip install -r requirements.txt
    python app.py            # http://localhost:8080

## Deploy on Railway
Push to GitHub, then New Project -> Deploy from GitHub repo. Railway installs
requirements.txt and starts from the Procfile
(`gunicorn app:app --bind 0.0.0.0:$PORT`). No env vars needed.

## Files
| File | Purpose |
|------|---------|
| `app.py` | Flask app: `/`, `/api/align`, `/api/download/<id>`. |
| `normalizer.py` | The brand engine. |
| `templates/index.html` | The page. |
| `fonts/Barlow-*.ttf` | Embedded into every processed deck — **must stay in the repo**. |
| `xelix_blue.png`, `xelix_white.png` | Wordmarks stamped on slides. |
| `requirements.txt`, `Procfile` | Deploy config. |

> The `fonts/` TTFs and the logo PNGs are required at runtime — keep them committed.
