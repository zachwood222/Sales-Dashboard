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
  const headers = rawDE[1]; // row index 1 is the real header
  const dailyRows = rawDE.slice(2).filter(r => r[0] && r[1]);

  const daily = dailyRows.map(r => ({
    date: normalizeDate(r[0]),
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
    yearMonth: normalizeMonth(r[19] || r[0]),
  }));

  // ── Monthly Summary ──
  const rawMS = xlsx.utils.sheet_to_json(wb.Sheets['Monthly Summary'], { header: 1 });
  const monthly = rawMS.slice(2)
    .filter(r => r[0] instanceof Date || typeof r[0] === 'number')
    .map(r => {
      const d = r[0] instanceof Date ? r[0] : new Date(Math.round((r[0] - 25569) * 86400 * 1000));
      return {
        month: normalizeMonth(d),
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

  const dailyByDate = [...daily].sort((a, b) => b.date.localeCompare(a.date));
  const monthlyByMonth = [...monthly].sort((a, b) => b.month.localeCompare(a.month));
  const rolling7Day = dailyByDate.slice(0, 7);
  const rolling30Day = dailyByDate.slice(0, 30);

  const dailyTotals = summarizeDaily(dailyByDate);
  const monthlyTotals = summarizeMonthly(monthlyByMonth);
  const monthlyAverages = averageMonthly(monthlyByMonth);

  const periods = buildPeriods(dailyByDate, monthlyByMonth);

  return {
    daily,
    monthly,
    kpis,
    weekdays,
    dailyByDate,
    monthlyByMonth,
    rolling: {
      days7: rolling7Day,
      days30: rolling30Day,
      summary7: summarizeDaily(rolling7Day),
      summary30: summarizeDaily(rolling30Day),
    },
    summaries: {
      dailyTotals,
      monthlyTotals,
      monthlyAverages,
    },
    periods,
  };
}

function xlsDateToISO(serial) {
  const d = new Date(Math.round((serial - 25569) * 86400 * 1000));
  return d.toISOString().slice(0, 10);
}

function normalizeDate(value) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'number') return xlsDateToISO(value);
  const date = new Date(String(value));
  if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  return String(value || '');
}

function normalizeMonth(value) {
  const isoDate = normalizeDate(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) return isoDate.slice(0, 7);
  const asMonth = String(value || '').trim();
  const mm = asMonth.match(/^(\d{4})-(\d{1,2})$/);
  if (mm) return `${mm[1]}-${String(mm[2]).padStart(2, '0')}`;
  return asMonth;
}

function summarizeDaily(rows) {
  return rows.reduce((acc, row) => {
    acc.subtotal += row.subtotal || 0;
    acc.taxes += (row.stateTax || 0) + (row.cityTax || 0);
    acc.deposits += row.deposits || 0;
    acc.grossMarginDollar += row.grossMarginDollar || 0;
    return acc;
  }, { subtotal: 0, taxes: 0, deposits: 0, grossMarginDollar: 0 });
}

function summarizeMonthly(rows) {
  return rows.reduce((acc, row) => {
    acc.sales += row.sales || 0;
    acc.subtotal += row.subtotal || 0;
    acc.grossMargin += row.grossMargin || 0;
    acc.salesPerSqFt += row.salesPerSqFt || 0;
    return acc;
  }, { sales: 0, subtotal: 0, grossMargin: 0, salesPerSqFt: 0 });
}

function averageMonthly(rows) {
  if (!rows.length) return { sales: 0, subtotal: 0, grossMargin: 0, salesPerSqFt: 0 };
  const totals = summarizeMonthly(rows);
  return {
    sales: totals.sales / rows.length,
    subtotal: totals.subtotal / rows.length,
    grossMargin: totals.grossMargin / rows.length,
    salesPerSqFt: totals.salesPerSqFt / rows.length,
  };
}

function buildPeriods(dailyRows, monthlyRows) {
  const today = dailyRows[0]?.date || null;
  const yesterday = dailyRows[1]?.date || null;
  const thisMonth = monthlyRows[0]?.month || null;
  const lastMonth = monthlyRows[1]?.month || null;
  const ytd = thisMonth ? `${thisMonth.slice(0, 4)}-01` : null;
  return { today, yesterday, thisMonth, lastMonth, ytd };
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
