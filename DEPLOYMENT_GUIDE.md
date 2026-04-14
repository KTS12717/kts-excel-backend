# KTS Fleet Manager — Complete Deployment Guide
## GitHub Pages (Frontend) + Render.com (Backend)

---

## What You Are Building

```
Your Computer                   The Internet
─────────────────                ──────────────────────────────────────────
                                 GitHub Repo 1: kts12717/kts-fleet-manager
  index.html  ────push──────▶   (GitHub Pages serves your app for free)
                                 https://kts12717.github.io/kts-fleet-manager

                                 GitHub Repo 2: kts12717/kts-excel-backend
  kts_excel_backend/ ─push──▶   (Render.com runs your Python server)
                                 https://kts-excel-backend.onrender.com
```

Both services are **completely free**.
GitHub Pages hosts the HTML app.
Render.com runs the Python backend that generates the Excel files.

---

## PART 1 — Set Up the Backend on Render.com

### Step 1.1 — Create a new GitHub repo for the backend

1. Go to https://github.com/new
2. Repository name: **kts-excel-backend**
3. Make it **Public** (Render free tier requires public repo)
4. Click **Create repository**
5. GitHub will show you a page with setup commands. Copy the repo URL.
   It will look like: `https://github.com/kts12717/kts-excel-backend.git`

### Step 1.2 — Upload the backend files to GitHub

You have two options:

**Option A — Upload through the GitHub website (no command line):**

1. Open the `kts_excel_backend` folder on your computer
2. Go to your new GitHub repo in the browser
3. Click **uploading an existing file**
4. Drag ALL these files and folders into the upload area:
   - `app.py`
   - `requirements.txt`
   - `Procfile`
   - `exporters/` folder (drag the whole folder)
   - `templates_excel/` folder (drag the whole folder)
5. Scroll down, click **Commit changes**

**Option B — Git command line (faster):**
```bash
cd kts_excel_backend
git init
git add .
git commit -m "Initial backend deployment"
git remote add origin https://github.com/kts12717/kts-excel-backend.git
git push -u origin main
```

### Step 1.3 — Create a free account on Render.com

1. Go to https://render.com
2. Click **Get Started for Free**
3. Sign up with your GitHub account (click "Sign in with GitHub")
4. Authorize Render to access your GitHub

### Step 1.4 — Deploy the backend on Render

1. On the Render dashboard, click **New +** → **Web Service**
2. Click **Connect a repository**
3. Find and select **kts-excel-backend**
4. Click **Connect**
5. Fill in these settings:

   | Field | Value |
   |-------|-------|
   | Name | kts-excel-backend |
   | Region | Oregon (US West) or closest to you |
   | Branch | main |
   | Runtime | **Python 3** |
   | Build Command | `pip install -r requirements.txt` |
   | Start Command | `gunicorn app:app --workers 2 --timeout 60 --bind 0.0.0.0:$PORT` |
   | Plan | **Free** |

6. Click **Advanced** and add this environment variable:

   | Key | Value |
   |-----|-------|
   | ALLOWED_ORIGIN | https://kts12717.github.io |

7. Click **Create Web Service**

Render will now build and deploy your backend. This takes about 2-3 minutes.
When it says **Live**, you will see your URL at the top of the page.
Copy it — it looks like: `https://kts-excel-backend.onrender.com`

### Step 1.5 — Test that the backend is working

Open this URL in your browser (replace with your actual URL):
```
https://kts-excel-backend.onrender.com/api/health
```

You should see:
```json
{
  "ha0935_template": true,
  "hv0713_template": true,
  "status": "ok"
}
```

If you see that, the backend is running correctly.

---

## PART 2 — Connect the Frontend to the Backend

### Step 2.1 — Edit index.html

Open `index.html` in any text editor (Notepad, VS Code, etc.)

Find this section near the top of the file (around line 30-45):
```html
<script>
  window.KTS_BACKEND_URL = "";  // ← paste your backend URL here
</script>
```

Change the empty string to your Render URL:
```html
<script>
  window.KTS_BACKEND_URL = "https://kts-excel-backend.onrender.com";
</script>
```

Save the file.

### Step 2.2 — Upload the updated index.html to GitHub Pages

1. Go to https://github.com/kts12717/kts-fleet-manager
2. Find `index.html` in the file list
3. Click on it, then click the **pencil icon** (Edit)
4. Select all the text (Ctrl+A), delete it
5. Open your new `index.html` in a text editor, select all (Ctrl+A), copy (Ctrl+C)
6. Paste into the GitHub editor (Ctrl+V)
7. Click **Commit changes**

**Or** use drag-and-drop: delete the old file, upload the new one.

---

## PART 3 — Verify Everything Works

1. Open your app: `https://kts12717.github.io/kts-fleet-manager`
2. Log in as admin
3. Go to the **Reports** tab
4. You should now see two types of download buttons:
   - **📋 Download HA0935 Official Template** — uses the backend
   - **📋 Download HV0713 Official Template** — uses the backend
   - **📊 Download KTS Client Tracker** — browser-only (always worked)
   - **📊 Download Priority Tracker** — browser-only (always worked)
5. Click **Download HA0935 Official Template**
6. A spinner will appear for 2-3 seconds, then the file downloads
7. Open it in Excel — it should look exactly like your original template

---

## IMPORTANT: Render Free Tier Behavior

The free tier on Render has one quirk:
**The server "sleeps" after 15 minutes of no activity.**

When someone clicks Download after the server has been idle:
- First request takes **30-60 seconds** (server waking up)
- All subsequent requests are fast (2-3 seconds)

To avoid this, you can either:
1. Accept it — users just wait a bit on the first download
2. Upgrade to Render's paid tier ($7/month) for always-on
3. Use **Railway.app** instead (similar free tier, slightly better cold start)

---

## Troubleshooting

**Problem: Health check shows `"ha0935_template": false`**
Solution: The template files didn't upload correctly to GitHub.
Check that `templates_excel/HA0935_template.xlsx` and `HV0713_template.xlsx`
are visible in your GitHub repo.

**Problem: Download button shows "Export failed: Network error"**
Solution: The backend URL is wrong or the server is sleeping.
1. Check `window.KTS_BACKEND_URL` is set correctly in index.html
2. Visit the health check URL directly to wake the server up
3. Try the download again

**Problem: "CORS error" in browser console**
Solution: The `ALLOWED_ORIGIN` environment variable on Render doesn't match
your GitHub Pages URL exactly.
Go to Render → your service → Environment → check ALLOWED_ORIGIN equals
`https://kts12717.github.io` (no trailing slash, no path).

**Problem: Render build fails**
Solution: Check the Render build logs. Common issue:
The `templates_excel/` folder was not included in the git push.
Make sure both `.xlsx` files are in the repo.

---

## Directory Structure Summary

**Repo 1: kts12717/kts-fleet-manager (GitHub Pages)**
```
index.html          ← your entire frontend app
```

**Repo 2: kts12717/kts-excel-backend (Render.com)**
```
app.py              ← Flask server, API routes
requirements.txt    ← Python dependencies
Procfile            ← tells Render how to start the app
exporters/
  __init__.py
  base.py           ← template loader (the safe-copy engine)
  ha0935.py         ← Khan Transportation exporter
  hv0713.py         ← Priority Transportation exporter
templates_excel/
  HA0935_template.xlsx    ← original master (never modified)
  HV0713_template.xlsx    ← original master (never modified)
```

---

## Total Cost

| Service | What it does | Cost |
|---------|-------------|------|
| GitHub Pages | Hosts your HTML app | Free forever |
| Render.com Web Service | Runs Python backend | Free (with sleep) |
| GitHub (backend repo) | Stores backend code | Free forever |
| **Total** | | **$0/month** |
