# Xelix Brand Aligner — web app

A single-purpose web tool: upload one Xelix deck, align it to the brand
guidelines, download the result. Wraps `normalizer.py` (no merge, no AI — that's
by design; assemble your combined deck first, then run it through here).

## Run locally

```bash
pip install -r requirements.txt
python app.py            # http://localhost:8080
```

## Deploy on Railway

1. Push this folder to a GitHub repo (see below).
2. In Railway: **New Project → Deploy from GitHub repo** → pick the repo.
3. Railway auto-detects Python, runs `pip install -r requirements.txt`, and
   starts the app from the `Procfile`:
   `web: gunicorn app:app --bind 0.0.0.0:$PORT --timeout 180 --workers 2`
   No environment variables are required (`$PORT` is provided by Railway).
4. Once it's live, open the generated URL and you're done.

The `--timeout 180` gives large (100+ slide) decks room to process.

## Add these files to git

From inside this folder:

```bash
git init
git add .
git commit -m "Xelix Brand Aligner web app"
git branch -M main
git remote add origin https://github.com/<you>/xelix-brand-aligner.git
git push -u origin main
```

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask app: `/`, `/api/align`, `/api/download/<id>`. |
| `normalizer.py` | The brand engine (fonts, wordmark, backgrounds, contrast). |
| `templates/index.html` | The page (Xelix-branded, drag-and-drop upload). |
| `xelix_blue.png`, `xelix_white.png` | Wordmarks stamped on slides. |
| `requirements.txt`, `Procfile` | Railway deploy config. |

## Two modes

- **Logo & fonts (default):** Barlow everywhere + wordmark bottom-right.
  Non-destructive — backgrounds/colours/layouts untouched.
- **Full brand alignment:** also imposes canonical white / hero-gradient
  backgrounds and brand text colours. For plain decks; it replaces backgrounds
  and recolours text.

## Notes

- Uploads are written to a temp dir for processing and the input is deleted
  immediately after; aligned outputs live in temp storage (cleared on restart),
  so download promptly. Railway's filesystem is ephemeral.
- Install the free Barlow Google Font wherever the deck is opened so it renders
  as intended.
