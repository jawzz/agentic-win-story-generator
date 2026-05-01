# UiPath Win Story Generator

Single-URL tool for generating branded UiPath agentic use case slides matching the Agent Use Case Template v10.

## What it does

1. Paste raw notes about an agentic automation win into the AI intake box
2. Backend calls Claude API to extract structured fields
3. Review/tweak the auto-filled form (variable step count, variable outcomes, optional impact metrics)
4. Hit Generate — downloads a single-slide branded `.pptx` in dark or light theme

## What's different from the original Win Story Generator

- Single-slide output (not multi-slide)
- Built specifically for agentic orchestration stories: AGENT / BOT / HUMAN actor roles on each step
- Supports variable step count (3-9 steps, auto-sizes tiles)
- Supports 1-5 measured outcome tiles
- New optional fields: **Problem metrics** (0-4 stats), **Attributable impact**, **Downstream impact**
- Embedded high-resolution UiPath logo (2400px from SVG source)
- Dark/Light theme toggle

## Local setup

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=sk-ant-...
python app.py
# Open http://localhost:8080
```

## Deploy to Render (recommended — same as the old generator)

1. Push this folder to a new GitHub repo:
   ```powershell
   cd "C:\Users\josh.fox\Documents\Claude\Projects\Agentic Win Story Generator"
   git init
   git branch -M main
   git add -A
   git commit -m "Initial commit: Agentic Win Story Generator"
   git remote add origin https://github.com/<your-user>/agentic-win-story-generator.git
   git push -u origin main
   ```

2. In Render dashboard:
   - **New → Web Service** → Connect your GitHub account → pick the repo
   - **Name**: `agentic-win-story-generator` (pre-filled from `render.yaml`)
   - **Runtime**: Python 3 (auto-detected)
   - **Build command**: `pip install -r requirements.txt`
   - **Start command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 45`
   - **Environment variable**: add `ANTHROPIC_API_KEY` with your key (the `sync: false` in render.yaml is intentional — forces you to set it per environment)
   - **Plan**: Free
   - Click **Create Web Service**

3. Wait ~2-3 minutes for build + deploy. Your URL will be `https://agentic-win-story-generator.onrender.com` (or similar).

4. The old Win Story Generator is untouched — it runs on its own Render service at its original URL.

## Env vars

| Variable | Required | Description |
|----------|----------|-------------|
| `ANTHROPIC_API_KEY` | Yes | Your Anthropic API key. AI extraction won't work without it. |
| `PORT` | No | Defaults to 8080 |

## Files

- `app.py` — Flask backend (routes: `/extract`, `/generate`, `/health`)
- `generate_pptx.py` — PPTX engine using python-pptx, builds the Agent Use Case Template
- `static/index.html` — Frontend (dark UI, dynamic form)
- `static/uipath-logo.svg` — SVG logo used by the web UI (crisp at any size)
- `static/uipath_logo_2400.png` — 2400px white logo embedded in generated pptx (dark theme)
- `static/uipath_logo_2400_dark.png` — 2400px dark logo for light theme
- `requirements.txt` — Python deps (flask, gunicorn, python-pptx, Pillow)
- `render.yaml` — Render blueprint
- `Procfile`, `Dockerfile` — alternate deploy targets

## Template layout

The generated slide (13.33" x 7.5") has these zones, top to bottom:

1. **Breadcrumb** — Industry / Function / Use case name, plus UiPath logo top-right
2. **Title + subtitle + company** — title flows to subtitle in a single adaptive text frame; short titles (< 50 chars) trigger a more compact top zone
3. **Problem card** (solid dark orange) + **Solution card** (solid dark teal) — side by side
   - Problem: description + 0-4 metric stats below
   - Solution: description + auto-wrapping white capability pills
4. **What the automation does** — gray container with numbered step tiles and chevron arrows. Adaptive sizing: 3-6 steps get wide tiles, 7-9 steps shrink to fit.
5. **Measured outcomes** — 1-5 solid orange tiles with big white stat numbers
6. **Attributable impact** (teal) + **Downstream impact** (gold) — side by side inline cards. Both optional.

Colors and fonts follow UiPath brand: primary orange `#FA4616`, teal `#0BA2B3`, gold `#DA9100`, Poppins throughout.
