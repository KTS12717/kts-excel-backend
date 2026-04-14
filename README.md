# KTS Excel Export Backend

Python Flask backend that generates exact-template Excel exports for:
- **HA0935** — Khan Transportation (Routes Khan1–Khan10)
- **HV0713** — Priority Transportation (Routes 6770–6771)

## Quick Deploy on Render.com (free tier)

1. Push this folder to a GitHub repo
2. Go to render.com → New → Web Service
3. Connect your repo
4. Settings:
   - **Runtime:** Python 3
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn app:app`
5. Add environment variable:
   - `ALLOWED_ORIGIN` = `https://kts12717.github.io`
6. Deploy → copy the URL (e.g. `https://kts-excel-backend.onrender.com`)
7. In your `index.html`, set:
   ```html
   <script>window.KTS_BACKEND_URL = "https://kts-excel-backend.onrender.com";</script>
   ```

## Local Development

```bash
pip install -r requirements.txt
python app.py
# → http://localhost:5000

# In index.html:
window.KTS_BACKEND_URL = "http://localhost:5000";
```

## API Endpoints

| Method | Endpoint              | Returns       |
|--------|----------------------|---------------|
| GET    | /api/health          | JSON status   |
| POST   | /api/export/ha0935   | .xlsx download |
| POST   | /api/export/hv0713   | .xlsx download |

## File Structure

```
kts_excel_backend/
├── app.py                     Flask server + API routes
├── requirements.txt
├── Procfile                   Render/Heroku deploy config
├── exporters/
│   ├── base.py                Template loader (shutil.copy + openpyxl)
│   ├── ha0935.py              HA0935 Khan Transportation exporter
│   └── hv0713.py              HV0713 Priority Transportation exporter
└── templates_excel/
    ├── HA0935_template.xlsx   Original master template (never modified)
    └── HV0713_template.xlsx   Original master template (never modified)
```

## How It Works

1. Frontend POSTs JSON with month date + route data
2. Backend copies the master `.xlsx` template to a temp file
3. openpyxl opens the copy, writes ONLY to yellow (#FFFFFF99) input cells
4. All formulas, borders, print settings, margins preserved exactly
5. Finished file returned as binary download
6. Temp file deleted

## Important

The `templates_excel/` folder must contain the original unmodified `.xlsx` files.
These are the master templates — they are **never modified** by the export process.
