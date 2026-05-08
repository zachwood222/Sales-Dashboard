const express = require('express');
const axios = require('axios');
const xlsx = require('xlsx');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');
const { mergeDailySources, buildSourceStats } = require('./src/data/mergeSources');

const app = express();
const PORT = process.env.PORT || 3000;
const FILE_URL = process.env.FILE_URL; // Google Drive or Dropbox direct-download URL
const LOCAL_FILE = path.join(__dirname, 'data', 'fowhand.xlsm');
const UPLOAD_STORE_FILE = path.join(__dirname, 'data', 'uploadedDaily.json');
const MERGE_POLICY = process.env.MERGE_POLICY || 'upload_wins';

let cachedData = null;
let lastUpdated = null;
let uploadStore = { records: [], lastUploadAt: null };

app.use(express.json({ limit: '2mb' }));

function loadUploadStore() {
  try {
    if (!fs.existsSync(UPLOAD_STORE_FILE)) return;
    const parsed = JSON.parse(fs.readFileSync(UPLOAD_STORE_FILE, 'utf8'));
    uploadStore = {
      records: Array.isArray(parsed.records) ? parsed.records : [],
      lastUploadAt: parsed.lastUploadAt || null,
    };
  } catch (err) {
    console.error('Failed to load upload store:', err.message);
  }
}

function saveUploadStore() {
  fs.mkdirSync(path.dirname(UPLOAD_STORE_FILE), { recursive: true });
  fs.writeFileSync(UPLOAD_STORE_FILE, JSON.stringify(uploadStore, null, 2));
}

// ── helpers ──────────────────────────────────────────────────────────────────

function toDirectDownload(url) {
  const gsheet = url.match(/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (gsheet) return `https://docs.google.com/spreadsheets/d/${gsheet[1]}/export?format=xlsx`;
  const gdrive = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (gdrive && url.includes('drive.google.com')) {
    return `https://drive.google.com/uc?export=download&id=${gdrive[1]}`;
  }
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

function normalizeDailyRow(row, source = 'upload') {
  const dateVal = row.date;
  return {
    date: dateVal instanceof Date ? dateVal.toISOString().slice(0, 10)
      : typeof dateVal === 'number' ? xlsDateToISO(dateVal) : String(dateVal || ''),
    subtotal: +row.subtotal || 0,
    cash: +row.cash || 0,
    card: +row.card || 0,
    deposits: +row.deposits || 0,
    deliveryFee: +row.deliveryFee || 0,
    stateTax: +row.stateTax || 0,
    cityTax: +row.cityTax || 0,
    grossMarginPct: +row.grossMarginPct || 0,
    grossMarginDollar: +row.grossMarginDollar || 0,
    subtotalWithDelivery: +row.subtotalWithDelivery || 0,
    salesPerSqFt: +row.salesPerSqFt || 0,
    weekday: row.weekday || '',
    yearMonth: row.yearMonth || '',
    ticketId: row.ticketId || null,
    storeId: row.storeId || null,
    source,
  };
}

function parseWorkbook() {
  const wb = xlsx.readFile(LOCAL_FILE, { type: 'file', cellDates: true });

  const rawDE = xlsx.utils.sheet_to_json(wb.Sheets['Daily Entry'], { header: 1 });
  const dailyRows = rawDE.slice(2).filter(r => r[0] && r[1]);

  const workbookDaily = dailyRows.map(r => normalizeDailyRow({
    date: r[0], subtotal: r[1], cash: r[2], card: r[3], deposits: r[4], deliveryFee: r[5], stateTax: r[6], cityTax: r[7],
    grossMarginPct: r[12], grossMarginDollar: r[13], subtotalWithDelivery: r[14], salesPerSqFt: r[15], weekday: r[17], yearMonth: String(r[19] || ''),
  }, 'google'));

  const mergedDaily = mergeDailySources(workbookDaily, uploadStore.records, { policy: MERGE_POLICY });
  const monthly = computeMonthlyFromDaily(mergedDaily);
  const kpis = computeKpisFromDaily(mergedDaily);

  const rawWA = xlsx.utils.sheet_to_json(wb.Sheets['Weekday Analysis'], { header: 1 });
  const weekdays = rawWA.slice(2).filter(r => r[0]).map(r => ({
    day: r[0],
    avgSales: +r[1] || 0,
    avgSubtotal: +r[2] || 0,
    avgGrossMargin: +r[3] || 0,
    daysLoaded: +r[4] || 0,
    goalHitPct: +r[5] || 0,
  }));

  return { daily: mergedDaily, monthly, kpis, weekdays, sourceStats: buildSourceStats(mergedDaily, uploadStore.lastUploadAt) };
}

function computeMonthlyFromDaily(daily) {
  const byMonth = new Map();
  for (const row of daily) {
    const month = String(row.date || '').slice(0, 7);
    if (!month) continue;
    if (!byMonth.has(month)) byMonth.set(month, { month, sales: 0, deliveryFees: 0, subtotal: 0, deposits: 0, grossMargin: 0, salesPerSqFt: 0, goalVariance: 0, _days: 0 });
    const item = byMonth.get(month);
    item.sales += row.subtotalWithDelivery || row.subtotal || 0;
    item.deliveryFees += row.deliveryFee || 0;
    item.subtotal += row.subtotal || 0;
    item.deposits += row.deposits || 0;
    item.grossMargin += row.grossMarginDollar || 0;
    item.salesPerSqFt += row.salesPerSqFt || 0;
    item._days += 1;
  }
  return Array.from(byMonth.values()).map(m => ({ ...m, salesPerSqFt: m._days ? m.salesPerSqFt / m._days : 0, goalVariance: 0 })).sort((a, b) => a.month.localeCompare(b.month));
}

function computeKpisFromDaily(daily) {
  const subtotal = daily.reduce((sum, row) => sum + (row.subtotal || 0), 0);
  const grossMargin = daily.reduce((sum, row) => sum + (row.grossMarginDollar || 0), 0);
  const deposits = daily.reduce((sum, row) => sum + (row.deposits || 0), 0);
  const bestDay = daily.reduce((best, row) => (!best || row.subtotal > best.subtotal ? row : best), null);
  return {
    ytdSubtotal: subtotal,
    mtdSubtotal: subtotal,
    mtdGrossMargin: grossMargin,
    depositPct: subtotal ? deposits / subtotal : 0,
    bestDaySubtotal: bestDay ? bestDay.subtotal : 0,
    bestDayDate: bestDay ? bestDay.date : null,
    loadedDays: daily.length,
  };
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

app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/data', (req, res) => {
  if (!cachedData) return res.status(503).json({ error: 'Data not loaded yet' });
  res.json({ data: cachedData, lastUpdated, sourceStats: cachedData.sourceStats });
});

app.post('/api/upload-daily', async (req, res) => {
  const rows = Array.isArray(req.body?.rows) ? req.body.rows : [];
  if (!rows.length) return res.status(400).json({ error: 'rows[] required' });
  const normalized = rows.map(r => normalizeDailyRow(r, 'upload')).filter(r => r.date);
  uploadStore.records = mergeDailySources(uploadStore.records, normalized, { policy: 'upload_wins' });
  uploadStore.lastUploadAt = new Date().toISOString();
  saveUploadStore();
  await refreshData();
  res.json({ ok: true, uploaded: normalized.length, lastUploadAt: uploadStore.lastUploadAt });
});

app.get('/api/refresh', async (req, res) => {
  await refreshData();
  res.json({ ok: true, lastUpdated });
});

app.get('/health', (req, res) => res.json({ ok: true }));

cron.schedule('0 6 * * *', () => {
  console.log('Cron triggered — refreshing data');
  refreshData();
});

app.listen(PORT, async () => {
  console.log(`Server listening on port ${PORT}`);
  loadUploadStore();
  await refreshData();
});
