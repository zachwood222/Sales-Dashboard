const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const { parseDailyRecordFromText, normalizeCurrency, normalizeDate } = require('../src/ingest/parseTicketPdf');

const fixturePath = path.join(__dirname, 'fixtures', 'ocrSamples.json');
const samples = JSON.parse(fs.readFileSync(fixturePath, 'utf8'));

test('normalizeCurrency handles commas and missing decimals deterministically', () => {
  assert.deepEqual(normalizeCurrency('$1,234.50'), { value: 1234.5, confidence: 0.98, error: null });
  assert.deepEqual(normalizeCurrency('1234'), { value: 1234, confidence: 0.78, error: null });
});

test('normalizeDate marks ambiguous slash dates with reduced confidence', () => {
  const parsed = normalizeDate('04/05/2026');
  assert.equal(parsed.value, '2026-04-05');
  assert.equal(parsed.error, 'ambiguous_date');
  assert.equal(parsed.confidence, 0.55);
});

test('fixture: commas_and_currency_symbols parses all fields', () => {
  const result = parseDailyRecordFromText(samples[0].text);
  assert.equal(result.record.subtotal, 1234.5);
  assert.equal(result.record.deposits, 50);
  assert.equal(result.record.cityTax, 25.12);
  assert.equal(result.metadata.date.value, '2026-03-14');
});

test('fixture: missing decimals defaults missing values to zero', () => {
  const result = parseDailyRecordFromText(samples[1].text);
  assert.equal(result.record.deliveryFee, 0);
  assert.equal(result.metadata.deliveryFee.error, 'missing');
  assert.equal(result.record.subtotal, 1234);
});

test('fixture: ambiguous_date preserves parse and flag', () => {
  const result = parseDailyRecordFromText(samples[2].text);
  assert.equal(result.record.date, '2026-04-05');
  assert.equal(result.metadata.date.error, 'ambiguous_date');
});


test('fixture: fowhand handwritten ticket style parses total sale as subtotal', () => {
  const result = parseDailyRecordFromText(samples[3].text);
  assert.equal(result.record.date, '2026-05-07');
  assert.equal(result.record.subtotal, 270.18);
  assert.equal(result.record.stateTax, 17.68);
});
