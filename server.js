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
  const headers = rawDE[1] || []; // row index 1 is the real header
  const dailyRows = rawDE.slice(2).filter(r => r[0] && r[1]);

  const normalizedHeaderIndex = new Map();
  headers.forEach((h, idx) => {
    if (!h) return;
    const key = String(h).trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ');
    normalizedHeaderIndex.set(key, idx);
  });

  function getColIndex(headerNames, fallbackIndex) {
    for (const name of headerNames) {
      const key = name.trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ');
      if (normalizedHeaderIndex.has(key)) return normalizedHeaderIndex.get(key);
    }
    return fallbackIndex;
  }

  function getNum(row, idx) {
    if (idx == null || idx < 0) return 0;
    return +row[idx] || 0;
  }

  const dailyCol = {
    subtotal: getColIndex(['Subtotal'], 1),
    cash: getColIndex(['Cash Sales', 'Cash'], 2),
    card: getColIndex(['Card Sales', 'Card'], 3),
    check: getColIndex(['Check Sales', 'Check'], 4),
    deposits: getColIndex(['Deposits', 'Deposit'], 4),
    deliveryFee: getColIndex(['Delivery Fee', 'Delivery Fees'], 5),
    stateTax: getColIndex(['State Tax'], 6),
    cityTax: getColIndex(['City Tax'], 7),
    grossMarginPct: getColIndex(['Gross Margin %', 'Gross Margin Pct'], 12),
    grossMarginDollar: getColIndex(['Gross Margin $', 'Gross Margin Dollar'], 13),
    subtotalWithDelivery: getColIndex(['Subtotal + Delivery', 'Subtotal With Delivery'], 14),
    salesPerSqFt: getColIndex(['Sales / Sq Ft', 'Sales Per Sq Ft'], 15),
    weekday: getColIndex(['Weekday', 'Day of Week'], 17),
    yearMonth: getColIndex(['Year Month', 'Year/Month'], 19),
  };

  const daily = dailyRows.map(r => {
    const cash = getNum(r, dailyCol.cash);
    const card = getNum(r, dailyCol.card);
    const check = getNum(r, dailyCol.check);
    const tenderTotal = cash + card + check;

    return {
      date: r[0] instanceof Date ? r[0].toISOString().slice(0, 10)
           : typeof r[0] === 'number' ? xlsDateToISO(r[0]) : String(r[0]),
      subtotal: getNum(r, dailyCol.subtotal),
      cash,
      card,
      check,
      tenderTotal,
      cashPct: tenderTotal > 0 ? cash / tenderTotal : 0,
      cardPct: tenderTotal > 0 ? card / tenderTotal : 0,
      checkPct: tenderTotal > 0 ? check / tenderTotal : 0,
      deposits: getNum(r, dailyCol.deposits),
      deliveryFee: getNum(r, dailyCol.deliveryFee),
      stateTax: getNum(r, dailyCol.stateTax),
      cityTax: getNum(r, dailyCol.cityTax),
      grossMarginPct: getNum(r, dailyCol.grossMarginPct),
      grossMarginDollar: getNum(r, dailyCol.grossMarginDollar),
      subtotalWithDelivery: getNum(r, dailyCol.subtotalWithDelivery),
      salesPerSqFt: getNum(r, dailyCol.salesPerSqFt),
      weekday: r[dailyCol.weekday] || '',
      yearMonth: String(r[dailyCol.yearMonth] || ''),
    };
  });

  const tenderSummary = daily.reduce((acc, row) => {
    acc.totalCash += row.cash;
    acc.totalCard += row.card;
    acc.totalCheck += row.check;
    return acc;
  }, { totalCash: 0, totalCard: 0, totalCheck: 0 });

  tenderSummary.tenderTotal = tenderSummary.totalCash + tenderSummary.totalCard + tenderSummary.totalCheck;
  tenderSummary.cashPct = tenderSummary.tenderTotal > 0 ? tenderSummary.totalCash / tenderSummary.tenderTotal : 0;
  tenderSummary.cardPct = tenderSummary.tenderTotal > 0 ? tenderSummary.totalCard / tenderSummary.tenderTotal : 0;
  tenderSummary.checkPct = tenderSummary.tenderTotal > 0 ? tenderSummary.totalCheck / tenderSummary.tenderTotal : 0;

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

  return { daily, monthly, kpis, weekdays, summary: { tender: tenderSummary } };
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
