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
let lastGoogleSync = null;
let lastUploadIngest = null;
let refreshInFlight = null;
let lastRefreshAtMs = 0;
const MIN_REFRESH_INTERVAL_MS = Number(process.env.MIN_REFRESH_INTERVAL_MS || 5 * 60 * 1000);

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

function pickBestHeaderRow(rows = [], maxRows = 5, expectedAliases = []) {
  let bestIdx = 0;
  let bestScore = -1;
  for (let i = 0; i < Math.min(maxRows, rows.length); i += 1) {
    const idxMap = buildHeaderIndexMap(rows[i] || []);
    const score = expectedAliases.reduce((count, aliases) => (
      resolveHeaderIndex(idxMap, aliases) === undefined ? count : count + 1
    ), 0);
    if (score > bestScore) {
      bestIdx = i;
      bestScore = score;
    }
  }
  return bestIdx;
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
  const dailyExpectedAliases = [
    ['Date'],
    ['Cash Sales', 'Cash'],
    ['Card Sales', 'Card'],
    ['Check Sales', 'Check', 'Deposits'],
    ['Subtotal'],
    ['State Tax'],
    ['City Tax'],
  ];
  const deHeaderRowIdx = pickBestHeaderRow(rawDE, 6, dailyExpectedAliases);
  const headers = rawDE[deHeaderRowIdx] || [];
  const deIdxMap = buildHeaderIndexMap(headers);
  logMissingHeaders('Daily Entry', deIdxMap, dailyExpectedAliases.map(aliases => ({ aliases })));

  const dailyRows = rawDE.slice(deHeaderRowIdx + 1).filter(r => r[0] && r[1]);
  const workbookDaily = dailyRows.map(r => normalizeDailyRow({
    date: r[0],
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
  }, 'google'));

  const mergedDaily = mergeDailySources(workbookDaily, uploadStore.records, { policy: MERGE_POLICY });

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
        month: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`,
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
    squareFeet: cell('B2'),
    avgInventoryCost: cell('B3'),
    dailyGoal: cell('B4'),
    monthlyGoal: cell('B5'),
    currentMonth: cell('B6'),
    mtdSubtotal: cell('E2'),
    mtdGoalVariance: cell('I2'),
    mtdGrossMargin: cell('M2'),
    mtdGMROI: cell('Q2'),
    ytdSubtotal: cell('E3'),
    salesPerSqFt: cell('I3'),
    grossMarginPerSqFt: cell('M3'),
    ytdGMROI: cell('Q3'),
    avgTicket: cell('E4'),
    depositPct: cell('I4'),
    bestDaySubtotal: cell('M4'),
    loadedDays: cell('Q4'),
    goalHitRate: cell('I5'),
    avgDailySubtotal: cell('M5'),
    bestDayDate: cell('B11'),
    bestDayGrossMargin: cell('B13'),
    bestDayWeekday: cell('B14'),
  };

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

async function refreshData(options = {}) {
  const { force = false, reason = 'unknown' } = options;
  const now = Date.now();

  if (refreshInFlight) return refreshInFlight;
  if (!force && lastRefreshAtMs && now - lastRefreshAtMs < MIN_REFRESH_INTERVAL_MS) {
    return;
  }

  refreshInFlight = (async () => {
    try {
      await downloadFile();
      cachedData = parseWorkbook();
      lastUpdated = new Date().toISOString();
      lastGoogleSync = lastUpdated;
      lastRefreshAtMs = Date.now();
      console.log(`Data refreshed at ${lastUpdated} (reason: ${reason})`);
    } catch (err) {
      console.error('Refresh failed:', err.message);
    } finally {
      refreshInFlight = null;
    }
  })();

  return refreshInFlight;
}

app.use(express.static(path.join(__dirname, 'public')));
app.use('/api/upload-ticket', express.raw({ type: ['application/pdf'], limit: '15mb' }));

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
  await refreshData({ force: true, reason: 'upload_daily' });
  res.json({ ok: true, uploaded: normalized.length, lastUploadAt: uploadStore.lastUploadAt });
});

app.get('/api/refresh', async (req, res) => {
  await refreshData({ force: true, reason: 'api_refresh' });
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
  refreshData({ force: true, reason: 'cron' });
});

app.listen(PORT, async () => {
  console.log(`Server listening on port ${PORT}`);
  loadUploadStore();
  await refreshData({ force: true, reason: 'startup' });
});
