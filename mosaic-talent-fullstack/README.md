# Mosaic Talent — AI-Powered Recruitment Platform

Full-stack web application: Flask backend + SQLite database + Gemini AI frontend.

## Stack
- **Backend**: Python/Flask 3.x
- **Database**: SQLite (via Python's built-in sqlite3)
- **Frontend**: Single-page HTML/CSS/JS (no framework, no build step)
- **AI**: Google Gemini 2.0 Flash (direct API calls from browser)
- **Excel Export**: openpyxl
- **Production Server**: Gunicorn

## Local Development

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run
python app.py

# 3. Open http://localhost:5000
```

## Deploy to Railway (Recommended — Free Tier)

1. Push this folder to a GitHub repo
2. Go to https://railway.app → New Project → Deploy from GitHub
3. Railway auto-detects the Procfile and runs `gunicorn app:app`
4. Done — your app is live!

## Deploy to Render

1. Push to GitHub
2. Go to https://render.com → New Web Service → Connect repo
3. Build Command: `pip install -r requirements.txt`
4. Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT`
5. Free tier available

## Deploy to Heroku

```bash
heroku create mosaic-talent-yourname
git push heroku main
heroku open
```

## Deploy to Fly.io

```bash
fly launch
fly deploy
```

## Key Features
- ✦ **JD Parser**: Upload or paste JDs → Gemini scans for keywords, maps to Bloom's Taxonomy
- 🔀 **Hybrid Role Detection**: Auto-detects cross-department roles and mixes questions
- ◈ **Assessment Generator**: 11 tailored questions with difficulty levels
- 👥 **Candidate Portal**: Timed tests with progress tracking
- 📊 **Leaderboard**: AI-scored rankings with integrity detection
- 📥 **Excel Export**: 3-sheet workbook (Summary, Responses, Analytics)
- 🗃️ **SQLite Database**: All data persisted across sessions

## API Endpoints
- `GET  /api/stats` — Dashboard stats
- `GET  /api/jobs` — All jobs
- `POST /api/jobs` — Create job
- `GET  /api/assessments/:id` — Get assessment
- `POST /api/assessments` — Create assessment
- `GET  /api/candidates` — All candidates (with job/assessment data)
- `GET  /api/candidates/:id` — Single candidate with questions+answers
- `POST /api/candidates` — Register candidate
- `PUT  /api/candidates/:id` — Update candidate (submit evaluation)
- `DELETE /api/candidates/:id` — Remove candidate
- `GET  /api/logs` — System audit log
- `GET  /api/export/excel` — Download Excel workbook

## Architecture Notes
The Gemini API key is embedded in the frontend HTML. In production:
1. Move the key to an environment variable
2. Proxy Gemini calls through the Flask backend `/api/ai/*`
3. This prevents key exposure in browser source

## Database
SQLite file is stored at `instance/mosaic.db`. For production with multiple workers, 
consider upgrading to PostgreSQL (replace sqlite3 with psycopg2 + adapt queries).
