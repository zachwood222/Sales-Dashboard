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
