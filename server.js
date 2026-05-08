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

  const computedKpis = computeKpis({ daily, monthly });

  return {
    daily,
    monthly,
    kpis: {
      source: 'sheet',
      values: kpis,
    },
    computedKpis,
    weekdays,
  };
}

function computeKpis({ daily, monthly }) {
  const parsedDaily = daily
    .map(d => {
      const dateObj = new Date(`${d.date}T00:00:00Z`);
      return Number.isNaN(dateObj.getTime()) ? null : { ...d, dateObj };
    })
    .filter(Boolean)
    .sort((a, b) => a.dateObj - b.dateObj);

  if (!parsedDaily.length) {
    return {
      source: 'computed',
      mtd: null,
      bestDay: null,
      paymentMix: null,
      trends: null,
    };
  }

  const latestDay = parsedDaily[parsedDaily.length - 1];
  const year = latestDay.dateObj.getUTCFullYear();
  const month = latestDay.dateObj.getUTCMonth();
  const dayOfMonth = latestDay.dateObj.getUTCDate();

  const isSameMonth = d => d.dateObj.getUTCFullYear() === year && d.dateObj.getUTCMonth() === month;
  const mtdRows = parsedDaily.filter(isSameMonth);
  const mtdSales = sumBy(mtdRows, 'subtotalWithDelivery');
  const mtdSubtotal = sumBy(mtdRows, 'subtotal');
  const mtdGrossMarginDollar = sumBy(mtdRows, 'grossMarginDollar');
  const mtdGrossMarginPct = mtdSubtotal > 0 ? mtdGrossMarginDollar / mtdSubtotal : null;
  const avgDailySubtotal = mtdRows.length > 0 ? mtdSubtotal / mtdRows.length : null;

  const bestDayRow = mtdRows.reduce((best, row) => {
    if (!best || row.subtotal > best.subtotal) return row;
    return best;
  }, null);

  const paymentCash = sumBy(mtdRows, 'cash');
  const paymentCard = sumBy(mtdRows, 'card');
  const paymentCheck = sumBy(mtdRows, 'deposits');
  const paymentTotal = paymentCash + paymentCard + paymentCheck;

  const previousDay = parsedDaily.length > 1 ? parsedDaily[parsedDaily.length - 2] : null;
  const sameWeekdayLastWeek = parsedDaily.find(
    d => d.dateObj.getTime() === latestDay.dateObj.getTime() - (7 * 24 * 60 * 60 * 1000),
  ) || null;

  const priorMonthDate = new Date(Date.UTC(year, month - 1, dayOfMonth));
  const priorMonthId = `${priorMonthDate.getUTCFullYear()}-${String(priorMonthDate.getUTCMonth() + 1).padStart(2, '0')}`;
  const priorMonthRow = monthly.find(m => m.month === priorMonthId) || null;
  const priorMonthCutoffValue = priorMonthRow
    ? ((priorMonthRow.subtotal || 0) / daysInMonth(priorMonthDate.getUTCFullYear(), priorMonthDate.getUTCMonth() + 1)) * dayOfMonth
    : null;

  return {
    source: 'computed',
    mtd: {
      sales: mtdSales,
      subtotal: mtdSubtotal,
      grossMarginDollar: mtdGrossMarginDollar,
      grossMarginPct: mtdGrossMarginPct,
      avgDailySubtotal,
    },
    bestDay: bestDayRow
      ? {
          date: bestDayRow.date,
          subtotal: bestDayRow.subtotal,
          grossMarginDollar: bestDayRow.grossMarginDollar,
        }
      : null,
    paymentMix: {
      cash: { total: paymentCash, pct: paymentTotal > 0 ? paymentCash / paymentTotal : null },
      check: { total: paymentCheck, pct: paymentTotal > 0 ? paymentCheck / paymentTotal : null },
      card: { total: paymentCard, pct: paymentTotal > 0 ? paymentCard / paymentTotal : null },
      total: paymentTotal,
    },
    trends: {
      vsPreviousDay: buildSubtotalTrend(latestDay, previousDay),
      vsSameWeekdayLastWeek: buildSubtotalTrend(latestDay, sameWeekdayLastWeek),
      mtdVsPriorMonthSameDayCutoff: {
        currentMtdSubtotal: mtdSubtotal,
        priorMonthSameDayCutoffSubtotal: priorMonthCutoffValue,
        delta: priorMonthCutoffValue === null ? null : mtdSubtotal - priorMonthCutoffValue,
        deltaPct: priorMonthCutoffValue > 0 ? (mtdSubtotal - priorMonthCutoffValue) / priorMonthCutoffValue : null,
      },
    },
  };
}

function buildSubtotalTrend(current, baseline) {
  if (!current || !baseline) return null;
  const delta = current.subtotal - baseline.subtotal;
  return {
    currentDate: current.date,
    baselineDate: baseline.date,
    currentSubtotal: current.subtotal,
    baselineSubtotal: baseline.subtotal,
    delta,
    deltaPct: baseline.subtotal > 0 ? delta / baseline.subtotal : null,
  };
}

function sumBy(rows, key) {
  return rows.reduce((sum, row) => sum + (+row[key] || 0), 0);
}

function daysInMonth(year, monthOneBased) {
  return new Date(Date.UTC(year, monthOneBased, 0)).getUTCDate();
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
