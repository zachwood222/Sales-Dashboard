const express = require('express');
const axios = require('axios');
const xlsx = require('xlsx');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const { uploadMiddleware, ingestTicketUpload, MAX_FILE_SIZE_BYTES } = require('./src/ingest/ticketUpload');

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
  const headers = rawDE[1] || []; // row index 1 is the real header
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

  const computedKpis = computeKpis({ daily, monthly });
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


app.post('/api/upload-ticket', (req, res) => {
  uploadMiddleware(req, res, err => {
    if (err) {
      if (err instanceof multer.MulterError) {
        if (err.code === 'LIMIT_FILE_SIZE') {
          return res.status(413).json({
            error: {
              code: 'FILE_TOO_LARGE',
              message: `File exceeds size limit of ${MAX_FILE_SIZE_BYTES} bytes`,
            },
          });
        }

        return res.status(400).json({
          error: {
            code: 'MALFORMED_MULTIPART',
            message: err.message,
          },
        });
      }

      if (err.code === 'UNSUPPORTED_FILE_TYPE') {
        return res.status(415).json({
          error: {
            code: err.code,
            message: err.message,
            acceptedTypes: ['application/pdf'],
          },
        });
      }

      return res.status(400).json({
        error: {
          code: 'MALFORMED_PAYLOAD',
          message: err.message || 'Unable to process upload payload.',
        },
      });
    }

    if (!req.file) {
      return res.status(400).json({
        error: {
          code: 'MISSING_FILE',
          message: "Expected 'ticket' file field in multipart/form-data payload.",
        },
      });
    }

    const result = ingestTicketUpload(req.file);
    return res.status(202).json({ upload: result });
  });
});


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
