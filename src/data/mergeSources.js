function rowKey(row) {
  const base = row.date || '';
  const ticket = row.ticketId ? `|ticket:${row.ticketId}` : '';
  const store = row.storeId ? `|store:${row.storeId}` : '';
  return `${base}${ticket}${store}`;
}

function pickPreferred(existing, incoming, policy) {
  if (!existing) return incoming;
  if (policy === 'google_wins') return existing;
  if (policy === 'upload_wins') return incoming;
  return incoming;
}

function mergeDailySources(workbookRows = [], uploadRows = [], opts = {}) {
  const policy = opts.policy || 'upload_wins';
  const merged = new Map();

  for (const row of workbookRows) {
    const keyed = { ...row, source: row.source || 'google' };
    merged.set(rowKey(keyed), keyed);
  }

  for (const row of uploadRows) {
    const keyed = { ...row, source: row.source || 'upload' };
    const key = rowKey(keyed);
    const existing = merged.get(key);
    const selected = pickPreferred(existing, keyed, policy);
    if (existing && selected === keyed) {
      merged.set(key, { ...selected, source: 'merged', sourceDetail: 'upload_override' });
    } else {
      merged.set(key, selected);
    }
  }

  return Array.from(merged.values()).sort((a, b) => String(a.date).localeCompare(String(b.date)));
}

function buildSourceStats(dailyRows = [], lastUploadAt = null) {
  const counts = dailyRows.reduce((acc, row) => {
    const src = row.source || 'unknown';
    acc[src] = (acc[src] || 0) + 1;
    return acc;
  }, {});

  return {
    counts,
    total: dailyRows.length,
    lastUploadAt,
  };
}

module.exports = {
  mergeDailySources,
  buildSourceStats,
};
