# Sales Adjustments Filter

A small web app that takes one or more QuickBooks-style invoice details Excel exports and returns a single merged `.xlsx` containing only the rows matching the predefined spec/adjustment criteria. Each row in the output is tagged with the source file it came from. The output also includes a `Criteria` tab documenting exactly what was applied and how many rows came from each source.

## Criteria

A row is kept if it matches **any** of the following. All comparisons are case-insensitive.

| # | Rule |
|---|------|
| 1 | Company is `YSM Tickets` AND TextTags contains `MFG` |
| 2 | Company is `YS Katz` AND TextTags contains `SPEC` |
| 3 | Company is `GK LLC` AND Performer/Team is `Los Angeles Dodgers` AND Account Email is `YKRAMER@YSKG.NET` |
| 4 | Company is `YSA`, `YSA 2`, or `YSA 3` AND TextTags contains `SCHMECK` |
| 5 | Company is `YS TL` AND TextTags contains `SPEC` |

## Input expectations

Each uploaded file should be an `.xlsx` invoice details export. The app reads the **first sheet** of each file regardless of its name (Others, Sheet1, etc.) and requires these columns:

- `Company`
- `Performer/Team`
- `Account Email`
- `TextTags`

All other columns are passed through to the output unchanged. When two or more files are uploaded, a `Source File` column is prepended so you can tell which file each row came from. Single-file uploads keep the original column layout.

The combined upload size limit is 200 MB. Adjust `MAX_CONTENT_LENGTH` in `app.py` if you need more.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Then open http://localhost:8000

## Deploy to Railway

1. Push this repo to GitHub.
2. In Railway, click **New Project → Deploy from GitHub repo** and pick this repo.
3. Railway will auto-detect Python via `runtime.txt` and `requirements.txt`, then use the start command in `Procfile` / `railway.json`.
4. No environment variables are required. Railway provides `$PORT` automatically.
5. Once deployed, click **Settings → Generate Domain** to get a public URL.

The health check endpoint is `/healthz`.

## Project layout

```
.
├── app.py              # Flask app + filter logic
├── templates/
│   └── index.html      # UI
├── requirements.txt
├── Procfile            # Start command for Railway/Heroku
├── runtime.txt         # Python version
├── railway.json        # Railway deploy config
├── .gitignore
└── README.md
```

## Updating the criteria

The criteria are defined in two places that must stay in sync:

- `CRITERIA` — a list of `{id, label, description}` dicts shown in the UI and Criteria tab.
- `apply_filters()` — the actual pandas masks.

Both live at the top of `app.py`.
