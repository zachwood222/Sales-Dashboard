const fs = require('fs/promises');

const FIELD_MAP = {
  subtotal: ['subtotal', 'sub total', 'sub-total'],
  cash: ['cash'],
  card: ['card', 'credit card', 'debit card'],
  deposits: ['deposits', 'deposit'],
  deliveryFee: ['delivery fee', 'delivery', 'del fee'],
  stateTax: ['state tax', 'st tax'],
  cityTax: ['city tax', 'cty tax']
};

function normalizeCurrency(rawValue) {
  if (rawValue == null) return { value: 0, confidence: 0, error: 'missing' };

  const token = String(rawValue).trim();
  if (!token) return { value: 0, confidence: 0, error: 'missing' };

  const cleaned = token
    .replace(/[$\s]/g, '')
    .replace(/[Oo](?=\d)/g, '0')
    .replace(/,/g, '')
    .replace(/\(([^)]+)\)/, '-$1');

  const numberMatch = cleaned.match(/-?\d+(?:\.\d{1,2})?/);
  if (!numberMatch) return { value: 0, confidence: 0, error: `invalid_currency:${token}` };

  const normalized = Number.parseFloat(numberMatch[0]);
  if (Number.isNaN(normalized)) return { value: 0, confidence: 0, error: `invalid_currency:${token}` };

  const hasDecimal = /\.\d{1,2}/.test(numberMatch[0]);
  const confidence = hasDecimal ? 0.98 : 0.78;
  return { value: normalized, confidence, error: null };
}

function normalizeDate(rawDate, options = {}) {
  const { preferMonthFirst = true } = options;
  if (rawDate == null || String(rawDate).trim() === '') {
    return { value: null, confidence: 0, error: 'missing' };
  }

  const src = String(rawDate).trim();
  const iso = src.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (iso) {
    const d = `${iso[1]}-${iso[2].padStart(2, '0')}-${iso[3].padStart(2, '0')}`;
    return { value: d, confidence: 0.99, error: null };
  }

  const slash = src.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (slash) {
    const a = Number(slash[1]);
    const b = Number(slash[2]);
    const year = slash[3].length === 2 ? `20${slash[3]}` : slash[3];

    let month = preferMonthFirst ? a : b;
    let day = preferMonthFirst ? b : a;
    let confidence = 0.93;
    let error = null;

    if (a <= 12 && b <= 12 && a !== b) {
      confidence = 0.55;
      error = 'ambiguous_date';
    } else if (a > 12) {
      month = b;
      day = a;
      confidence = 0.9;
    }

    const normalized = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    return { value: normalized, confidence, error };
  }

  return { value: null, confidence: 0, error: `invalid_date:${src}` };
}

function extractField(lines, aliases) {
  for (const line of lines) {
    const lower = line.toLowerCase();
    if (!aliases.some(alias => lower.includes(alias))) continue;

    const amount = line.match(/(-?\$?\(?\d[\d,]*(?:\.\d{1,2})?\)?)/);
    if (amount) return { raw: amount[1], sourceLine: line };
  }

  return { raw: null, sourceLine: null };
}

function toDailyRecord(parsed) {
  return {
    date: parsed.date.value,
    subtotal: parsed.subtotal.value,
    cash: parsed.cash.value,
    card: parsed.card.value,
    deposits: parsed.deposits.value,
    deliveryFee: parsed.deliveryFee.value,
    stateTax: parsed.stateTax.value,
    cityTax: parsed.cityTax.value
  };
}

function parseDailyRecordFromText(text, options = {}) {
  const lines = String(text || '')
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean);

  const dateLine = lines.find(line => /\bdate\b/i.test(line)) || lines.find(line => /\d{1,4}[/-]\d{1,2}[/-]\d{1,4}/.test(line));
  const dateToken = dateLine && dateLine.match(/(\d{1,4}[/-]\d{1,2}[/-]\d{1,4})/);

  const parsed = {
    date: normalizeDate(dateToken ? dateToken[1] : null, options)
  };

  for (const [field, aliases] of Object.entries(FIELD_MAP)) {
    const match = extractField(lines, aliases);
    parsed[field] = normalizeCurrency(match.raw);
    parsed[field].sourceLine = match.sourceLine;
  }

  return {
    record: toDailyRecord(parsed),
    metadata: parsed
  };
}

async function parseTicketPdf(pdfPath, opts = {}) {
  const { extractPdfText, ocrPage, preferMonthFirst = true } = opts;
  const raw = await fs.readFile(pdfPath);

  let pages = [];
  if (typeof extractPdfText === 'function') {
    pages = await extractPdfText(raw);
  }

  const parsedPages = [];
  for (let i = 0; i < pages.length; i += 1) {
    let text = pages[i] || '';
    if (!text.trim() && typeof ocrPage === 'function') {
      text = await ocrPage(raw, i);
    }

    const parsed = parseDailyRecordFromText(text, { preferMonthFirst });
    parsedPages.push(parsed);
  }

  return {
    records: parsedPages.map(p => ({ ...p.record, _meta: p.metadata })),
    parserMeta: {
      pageCount: pages.length,
      usedOcrOnPages: pages.map((p, i) => (!String(p || '').trim() ? i : null)).filter(v => v != null)
    }
  };
}

module.exports = {
  parseTicketPdf,
  parseDailyRecordFromText,
  normalizeCurrency,
  normalizeDate
};
