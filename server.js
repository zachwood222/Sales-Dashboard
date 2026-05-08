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


function normalizeHeaderName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildHeaderIndexMap(headers = []) {
  const map = {};
  headers.forEach((header, idx) => {
    const key = normalizeHeaderName(header);
    if (key && map[key] === undefined) map[key] = idx;
  });
  return map;
}

function resolveHeaderIndex(idxMap, names = []) {
  for (const name of names) {
    const idx = idxMap[normalizeHeaderName(name)];
    if (idx !== undefined) return idx;
  }
  return undefined;
}

function logMissingHeaders(sheetName, idxMap, expected) {
  const missing = expected.filter(({ aliases }) => resolveHeaderIndex(idxMap, aliases) === undefined);
  if (missing.length) {
    console.warn(`[parseWorkbook] ${sheetName}: missing headers -> ${missing.map(h => h.aliases[0]).join(', ')}. Using defaults.`);
  }
}

function getNum(row, idxMap, ...aliases) {
  const idx = resolveHeaderIndex(idxMap, aliases);
  return idx === undefined ? 0 : (+row[idx] || 0);
}

function getText(row, idxMap, ...aliases) {
  const idx = resolveHeaderIndex(idxMap, aliases);
  return idx === undefined ? '' : String(row[idx] || '');
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
  const deIdxMap = buildHeaderIndexMap(headers);
  logMissingHeaders('Daily Entry', deIdxMap, [
    { aliases: ['Cash Sales', 'Cash'] },
    { aliases: ['Card Sales', 'Card'] },
    { aliases: ['Check Sales', 'Check', 'Deposits'] },
    { aliases: ['Subtotal'] },
    { aliases: ['State Tax'] },
    { aliases: ['City Tax'] },
  ]);

  const dailyRows = rawDE.slice(2).filter(r => r[0] && r[1]);

  const daily = dailyRows.map(r => ({
    date: r[0] instanceof Date ? r[0].toISOString().slice(0, 10)
         : typeof r[0] === 'number' ? xlsDateToISO(r[0]) : String(r[0]),
    subtotal: getNum(r, deIdxMap, 'Subtotal'),
    cash: getNum(r, deIdxMap, 'Cash Sales', 'Cash'),
    card: getNum(r, deIdxMap, 'Card Sales', 'Card'),
    deposits: getNum(r, deIdxMap, 'Check Sales', 'Check', 'Deposits'),
    deliveryFee: getNum(r, deIdxMap, 'Delivery Fee', 'Delivery Fees'),
    stateTax: getNum(r, deIdxMap, 'State Tax'),
    cityTax: getNum(r, deIdxMap, 'City Tax'),
    grossMarginPct: getNum(r, deIdxMap, 'Gross Margin %', 'Gross Margin Pct'),
    grossMarginDollar: getNum(r, deIdxMap, 'Gross Margin $', 'Gross Margin Dollar'),
    subtotalWithDelivery: getNum(r, deIdxMap, 'Subtotal + Delivery', 'Subtotal With Delivery'),
    salesPerSqFt: getNum(r, deIdxMap, 'Sales / Sq Ft', 'Sales Per Sq Ft'),
    weekday: getText(r, deIdxMap, 'Weekday'),
    yearMonth: getText(r, deIdxMap, 'Year Month', 'Year-Month'),
  }));

  // ── Monthly Summary ──
  const rawMS = xlsx.utils.sheet_to_json(wb.Sheets['Monthly Summary'], { header: 1 });
  const msHeaders = rawMS[1] || rawMS[0] || [];
  const msIdxMap = buildHeaderIndexMap(msHeaders);
  logMissingHeaders('Monthly Summary', msIdxMap, [
    { aliases: ['Sales'] },
    { aliases: ['Delivery Fees', 'Delivery Fee'] },
    { aliases: ['Subtotal'] },
    { aliases: ['Deposits'] },
    { aliases: ['Gross Margin'] },
  ]);

  const monthly = rawMS.slice(2)
    .filter(r => r[0] instanceof Date || typeof r[0] === 'number')
    .map(r => {
      const d = r[0] instanceof Date ? r[0] : new Date(Math.round((r[0] - 25569) * 86400 * 1000));
      return {
        month: `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`,
        sales: getNum(r, msIdxMap, 'Sales'),
        deliveryFees: getNum(r, msIdxMap, 'Delivery Fees', 'Delivery Fee'),
        subtotal: getNum(r, msIdxMap, 'Subtotal'),
        deposits: getNum(r, msIdxMap, 'Deposits'),
        grossMargin: getNum(r, msIdxMap, 'Gross Margin', 'Gross Margin $'),
        salesPerSqFt: getNum(r, msIdxMap, 'Sales / Sq Ft', 'Sales Per Sq Ft'),
        goalVariance: getNum(r, msIdxMap, 'Goal Variance'),
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
  const waHeaders = rawWA[1] || rawWA[0] || [];
  const waIdxMap = buildHeaderIndexMap(waHeaders);
  logMissingHeaders('Weekday Analysis', waIdxMap, [
    { aliases: ['Day', 'Weekday'] },
    { aliases: ['Avg Sales', 'Average Sales'] },
    { aliases: ['Avg Subtotal', 'Average Subtotal'] },
    { aliases: ['Avg Gross Margin', 'Average Gross Margin'] },
    { aliases: ['Days Loaded'] },
    { aliases: ['Goal Hit %', 'Goal Hit Pct'] },
  ]);

  const weekdays = rawWA.slice(2).filter(r => r[0]).map(r => ({
    day: getText(r, waIdxMap, 'Day', 'Weekday') || r[0],
    avgSales: getNum(r, waIdxMap, 'Avg Sales', 'Average Sales'),
    avgSubtotal: getNum(r, waIdxMap, 'Avg Subtotal', 'Average Subtotal'),
    avgGrossMargin: getNum(r, waIdxMap, 'Avg Gross Margin', 'Average Gross Margin'),
    daysLoaded: getNum(r, waIdxMap, 'Days Loaded'),
    goalHitPct: getNum(r, waIdxMap, 'Goal Hit %', 'Goal Hit Pct'),
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
    console.log('Data refreshed at', lastUpdated);
  } catch (err) {
    console.error('Refresh failed:', err.message);
  }
}

// ── Routes ────────────────────────────────────────────────────────────────────

app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/data', (req, res) => {
  if (!cachedData) return res.status(503).json({ error: 'Data not loaded yet' });
  res.json({ data: cachedData, lastUpdated });
});

app.get('/api/refresh', async (req, res) => {
  await refreshData();
  res.json({ ok: true, lastUpdated });
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
