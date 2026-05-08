# Fowhand Dashboard — Deployment Guide

## What this does
A Node.js web server that:
- Downloads your `.xlsm` file daily from Google Drive or Dropbox
- Parses and serves the data as a live dashboard at your Render URL
- Refreshes automatically every day at 6 AM server time
- Has a manual "↻ Refresh" button on the dashboard

---

## Step 1 — Get your file's direct-download URL

### Google Drive
1. Upload your `.xlsm` to Google Drive
2. Right-click → **Share** → set to **"Anyone with the link"**
3. Copy the share link — it looks like:
   `https://drive.google.com/file/d/XXXXXXXXXXXXXXX/view?usp=sharing`
4. The server automatically converts this to a direct download link.

### Dropbox
1. Upload your `.xlsm` to Dropbox
2. Click **Share** → **Copy link**
3. The share link looks like:
   `https://www.dropbox.com/s/XXXXXXX/file.xlsm?dl=0`
4. The server automatically converts `?dl=0` → `?dl=1` for direct download.

---

## Step 2 — Push to GitHub

```bash
cd fowhand-dashboard
git init
git add .
git commit -m "Initial dashboard"
# Create a new repo on github.com, then:
git remote add origin https://github.com/YOUR_USERNAME/fowhand-dashboard.git
git push -u origin main
```

---

## Step 3 — Deploy on Render

1. Go to [render.com](https://render.com) and sign in (free account works)
2. Click **New → Web Service**
3. Connect your GitHub repo (`fowhand-dashboard`)
4. Render will auto-detect the `render.yaml` settings
5. Under **Environment Variables**, add:
   - Key: `FILE_URL`
   - Value: your Google Drive or Dropbox share URL from Step 1
6. Click **Create Web Service**

Render will build and deploy in ~2 minutes. Your dashboard will be live at:
`https://fowhand-dashboard.onrender.com`

---

## Step 4 — Updating the data

Whenever you save a new version of the `.xlsm` to the **same** Google Drive / Dropbox link:
- The dashboard will pick it up automatically **the next day at 6 AM**
- Or click the **↻ Refresh** button on the dashboard to pull it immediately

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Dashboard shows "Data not loaded yet" | Wait 30 sec after deploy, then click Refresh |
| Numbers look wrong | Make sure your share link has "Anyone with link" access |
| Dropbox link not working | Make sure the link ends in `?dl=1` or let the server convert it |
| File not found on Render | Check that `FILE_URL` env var is set in the Render dashboard |

---

## Local development

```bash
npm install
FILE_URL=your_share_url_here node server.js
# Visit http://localhost:3000
```

## Upload API: Ticket Ingestion

Use `POST /api/upload-ticket` with `multipart/form-data` and a `ticket` file field.

### Requirements
- Supported MIME type: `application/pdf`
- Max file size: `5 MB` (`5242880` bytes)

### Example request

```bash
curl -X POST http://localhost:3000/api/upload-ticket \
  -H "Accept: application/json" \
  -F "ticket=@./sample-ticket.pdf;type=application/pdf"
```

### Successful response (`202 Accepted`)

```json
{
  "upload": {
    "uploadId": "4f4ba347-0f9a-4bd0-bd76-4cf61d0663bf",
    "filename": "2026-05-08T14-10-00-111Z-sample-ticket.pdf",
    "ingestStatus": "queued",
    "parseErrors": []
  }
}
```

### Unsupported file type (`415 Unsupported Media Type`)

```json
{
  "error": {
    "code": "UNSUPPORTED_FILE_TYPE",
    "message": "Unsupported media type. Only application/pdf is accepted.",
    "acceptedTypes": ["application/pdf"]
  }
}
```

### Missing file field (`400 Bad Request`)

```json
{
  "error": {
    "code": "MISSING_FILE",
    "message": "Expected 'ticket' file field in multipart/form-data payload."
  }
}
```

### File too large (`413 Payload Too Large`)

```json
{
  "error": {
    "code": "FILE_TOO_LARGE",
    "message": "File exceeds size limit of 5242880 bytes"
  }
}
```
