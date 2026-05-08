const express = require('express');
const axios = require('axios');
const xlsx = require('xlsx');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const FILE_URL = process.env.FILE_URL; // Google Drive or Dropbox direct-download URL
const LOCAL_FILE = path.join(__dirname, 'data', 'fowhand.xlsm');

let cachedData = null;
let lastUpdated = null;
let lastGoogleSync = null;
let lastUploadIngest = null;

// ── helpers ──────────────────────────────────────────────────────────────────

function toDirectDownload(url) {
  // Google Sheets file links: export to XLSX
  const gsheet = url.match(/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (gsheet) return `https://docs.google.com/spreadsheets/d/${gsheet[1]}/export?format=xlsx`;

  // Google Drive file links: convert share link → direct download
  const gdrive = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (gdrive && url.includes('drive.google.com')) {
    return `https://drive.google.com/uc?export=download&id=${gdrive[1]}`;
  }

  // Dropbox: swap ?dl=0 → ?dl=1
  if (url.includes('dropbox.com')) return url.replace(/[?&]dl=0/, '').replace(/[?&]dl=1/, '') + '?dl=1';

  return url;
}

async function downloadFile() {
  if (!FILE_URL) {
    console.log('No FILE_URL set — using local file if present');
    return;
  }
  const directUrl = toDirectDownload(FILE_URL);
  console.log('Downloading file from:', directUrl);
  const response = await axios.get(directUrl, { responseType: 'arraybuffer', timeout: 30000 });
  fs.mkdirSync(path.dirname(LOCAL_FILE), { recursive: true });
  fs.writeFileSync(LOCAL_FILE, response.data);
  console.log('File downloaded:', LOCAL_FILE);
}

function parseWorkbook() {
  const wb = xlsx.readFile(LOCAL_FILE, { type: 'file', cellDates: true });

  // ── Daily Entry ──
  const rawDE = xlsx.utils.sheet_to_json(wb.Sheets['Daily Entry'], { header: 1 });
  const headers = rawDE[1]; // row index 1 is the real header
  const dailyRows = rawDE.slice(2).filter(r => r[0] && r[1]);

  const daily = dailyRows.map(r => ({
    date: r[0] instanceof Date ? r[0].toISOString().slice(0, 10)
         : typeof r[0] === 'number' ? xlsDateToISO(r[0]) : String(r[0]),
    subtotal:    +r[1]  || 0,
    cash:        +r[2]  || 0,
    card:        +r[3]  || 0,
    deposits:    +r[4]  || 0,
    deliveryFee: +r[5]  || 0,
    stateTax:    +r[6]  || 0,
    cityTax:     +r[7]  || 0,
    grossMarginPct: +r[12] || 0,
    grossMarginDollar: +r[13] || 0,
    subtotalWithDelivery: +r[14] || 0,
    salesPerSqFt: +r[15] || 0,
    weekday: r[17] || '',
    yearMonth: String(r[19] || ''),
  }));

  // ── Monthly Summary ──
  const rawMS = xlsx.utils.sheet_to_json(wb.Sheets['Monthly Summary'], { header: 1 });
  const monthly = rawMS.slice(2)
    .filter(r => r[0] instanceof Date || typeof r[0] === 'number')
    .map(r => {
      const d = r[0] instanceof Date ? r[0] : new Date(Math.round((r[0] - 25569) * 86400 * 1000));
      return {
        month: `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`,
        sales: +r[2] || 0,
        deliveryFees: +r[3] || 0,
        subtotal: +r[4] || 0,
        deposits: +r[5] || 0,
        grossMargin: +r[13] || 0,
        salesPerSqFt: +r[14] || 0,
        goalVariance: +r[15] || 0,
      };
    })
    .filter(m => m.sales > 0 || m.subtotal > 0);

  // ── Dashboard KPIs from Dashboard sheet ──
  const ds = wb.Sheets['Dashboard'];
  function cell(addr) { const c = ds[addr]; return c ? c.v : null; }

  const kpis = {
    squareFeet:      cell('B2'),
    avgInventoryCost:cell('B3'),
    dailyGoal:       cell('B4'),
    monthlyGoal:     cell('B5'),
    currentMonth:    cell('B6'),
    mtdSubtotal:     cell('E2'),
    mtdGoalVariance: cell('I2'),
    mtdGrossMargin:  cell('M2'),
    mtdGMROI:        cell('Q2'),
    ytdSubtotal:     cell('E3'),
    salesPerSqFt:    cell('I3'),
    grossMarginPerSqFt: cell('M3'),
    ytdGMROI:        cell('Q3'),
    avgTicket:       cell('E4'),
    depositPct:      cell('I4'),
    bestDaySubtotal: cell('M4'),
    loadedDays:      cell('Q4'),
    goalHitRate:     cell('I5'),
    avgDailySubtotal:cell('M5'),
    bestDayDate:     cell('B11'),
    bestDayGrossMargin: cell('B13'),
    bestDayWeekday:  cell('B14'),
  };

  // ── Weekday Analysis ──
  const rawWA = xlsx.utils.sheet_to_json(wb.Sheets['Weekday Analysis'], { header: 1 });
  const weekdays = rawWA.slice(2).filter(r => r[0]).map(r => ({
    day: r[0],
    avgSales: +r[1] || 0,
    avgSubtotal: +r[2] || 0,
    avgGrossMargin: +r[3] || 0,
    daysLoaded: +r[4] || 0,
    goalHitPct: +r[5] || 0,
  }));

  return { daily, monthly, kpis, weekdays };
}

function xlsDateToISO(serial) {
  const d = new Date(Math.round((serial - 25569) * 86400 * 1000));
  return d.toISOString().slice(0, 10);
}

async function refreshData() {
  try {
    await downloadFile();
    cachedData = parseWorkbook();
    lastUpdated = new Date().toISOString();
    lastGoogleSync = lastUpdated;
    console.log('Data refreshed at', lastUpdated);
  } catch (err) {
    console.error('Refresh failed:', err.message);
  }
}

// ── Routes ────────────────────────────────────────────────────────────────────

app.use(express.static(path.join(__dirname, 'public')));
app.use('/api/upload-ticket', express.raw({ type: ['application/pdf'], limit: '15mb' }));

app.get('/api/data', (req, res) => {
  if (!cachedData) return res.status(503).json({ error: 'Data not loaded yet' });
  res.json({ data: cachedData, lastUpdated, lastGoogleSync, lastUploadIngest });
});

app.get('/api/refresh', async (req, res) => {
  await refreshData();
  res.json({ ok: true, lastUpdated, lastGoogleSync, lastUploadIngest });
});

app.post('/api/upload-ticket', (req, res) => {
  try {
    const isPdf = (req.headers['content-type'] || '').includes('application/pdf');
    if (!isPdf || !req.body || !req.body.length) {
      return res.status(400).json({ error: 'Expected a PDF upload body.' });
    }
    lastUploadIngest = new Date().toISOString();
    return res.json({
      ok: true,
      ingestedAt: lastUploadIngest,
      lastUploadIngest,
      parsedTotals: { subtotal: 0, tax: 0, delivery: 0 },
      errors: ['PDF parser not configured on server yet; showing placeholder totals.']
    });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

app.get('/health', (req, res) => res.json({ ok: true }));

// ── Cron: refresh every day at 6 AM server time ───────────────────────────────
cron.schedule('0 6 * * *', () => {
  console.log('Cron triggered — refreshing data');
  refreshData();
});

// ── Boot ──────────────────────────────────────────────────────────────────────
app.listen(PORT, async () => {
  console.log(`Server listening on port ${PORT}`);
  await refreshData();
});
