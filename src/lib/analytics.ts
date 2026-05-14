import type { Mapping, Profile, Row } from './engine';
import { bucketHour, chooseHeader, makeTemplate, parseDateParts, safeText, normalize } from './engine';

export type RowAnalysis = {
  rowIndex: number;
  score: number;
  reasons: string[];
  invoiceNo: string;
  store: string;
  date: string;
  time: string;
  amount: number;
  sheet: string;
  row: Row;
};

export type NumericStat = { name: string; total: number; count: number };
export type StoreStat = { store: string; count: number; amount: number; suspicious: number };
export type TodStat = { bucket: string; count: number; amount: number };
export type Relation = { leftSheet: string; rightSheet: string; key: string; overlap: number; leftOnly: number; rightOnly: number };

const scoringProfiles: Record<Profile, Record<string, number>> = {
  balanced: { missing: 5, duplicate: 24, mismatch: 18, outlier: 16, relation: 10, oddHour: 4, blank: 3 },
  strict: { missing: 8, duplicate: 30, mismatch: 24, outlier: 22, relation: 14, oddHour: 6, blank: 5 },
  forensic: { missing: 10, duplicate: 36, mismatch: 28, outlier: 28, relation: 18, oddHour: 8, blank: 7 },
};

function median(nums: number[]): number {
  const a = [...nums].filter((n) => Number.isFinite(n)).sort((x, y) => x - y);
  if (!a.length) return 0;
  const mid = Math.floor(a.length / 2);
  return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function mad(nums: number[], med: number): number {
  const deviations = nums.map((n) => Math.abs(n - med));
  return median(deviations) || 1;
}

export function numericSummary(rows: Row[]): NumericStat[] {
  const map = new Map<string, { total: number; count: number }>();
  for (const row of rows) {
    for (const [k, v] of Object.entries(row)) {
      const n = typeof v === 'number' ? v : Number(String(v).replace(/,/g, ''));
      if (!Number.isFinite(n)) continue;
      const c = map.get(k) ?? { total: 0, count: 0 };
      c.total += n;
      c.count += 1;
      map.set(k, c);
    }
  }
  return Array.from(map.entries()).map(([name, val]) => ({ name, ...val })).sort((a, b) => Math.abs(b.total) - Math.abs(a.total)).slice(0, 8);
}

export function buildAnalysis(sheetName: string, headers: string[], rows: Row[], mapping: Mapping, profile: Profile) {
  const picked = {
    invoiceNo: chooseHeader(headers, mapping.invoiceNo),
    invoiceAmount: chooseHeader(headers, mapping.invoiceAmount),
    store: chooseHeader(headers, mapping.store),
    date: chooseHeader(headers, mapping.date),
    time: chooseHeader(headers, mapping.time),
    sku: chooseHeader(headers, mapping.sku),
    item: chooseHeader(headers, mapping.item),
    qty: chooseHeader(headers, mapping.qty),
    unitPrice: chooseHeader(headers, mapping.unitPrice),
    lineTotal: chooseHeader(headers, mapping.lineTotal),
    customer: chooseHeader(headers, mapping.customer),
    category: chooseHeader(headers, mapping.category),
  };

  const weights = scoringProfiles[profile];
  const invoiceMap = new Map<string, { count: number; amounts: number[] }>();
  const amounts: number[] = [];
  const analyses: RowAnalysis[] = [];

  rows.forEach((row, idx) => {
    const invoiceNo = safeText(picked.invoiceNo ? row[picked.invoiceNo] : '');
    const store = safeText(picked.store ? row[picked.store] : '');
    const rawDate = safeText(picked.date ? row[picked.date] : '');
    const rawTime = safeText(picked.time ? row[picked.time] : '');
    const amountRaw = picked.invoiceAmount ? row[picked.invoiceAmount] : (picked.lineTotal ? row[picked.lineTotal] : '');
    const amount = typeof amountRaw === 'number' ? amountRaw : Number(String(amountRaw).replace(/,/g, ''));
    const parsedTime = parseDateParts(rawTime || rawDate);
    const hour = parsedTime.hour;
    if (Number.isFinite(amount)) amounts.push(amount);
    if (invoiceNo) {
      const g = invoiceMap.get(invoiceNo) ?? { count: 0, amounts: [] };
      g.count += 1;
      if (Number.isFinite(amount)) g.amounts.push(amount);
      invoiceMap.set(invoiceNo, g);
    }
    const reasons: string[] = [];
    let score = 0;
    (['invoiceNo', 'invoiceAmount', 'store', 'date'] as const).forEach((f) => {
      if (!picked[f] || !safeText(row[picked[f] as string])) {
        score += weights.missing;
        reasons.push(`missing ${f}`);
      }
    });
    if (!invoiceNo && picked.invoiceNo) {
      score += weights.blank;
      reasons.push('blank invoice');
    }
    if (hour !== null && (hour < 6 || hour >= 22)) {
      score += weights.oddHour;
      reasons.push('odd hour');
    }
    if (!Number.isFinite(amount) || amount <= 0) {
      score += weights.blank;
      reasons.push('invalid amount');
    }
    analyses.push({
      rowIndex: idx,
      score,
      reasons,
      invoiceNo,
      store,
      date: parsedTime.date || rawDate,
      time: parsedTime.time || rawTime,
      amount: Number.isFinite(amount) ? amount : 0,
      sheet: sheetName,
      row,
    });
  });

  const med = median(amounts);
  const spread = mad(amounts, med);
  analyses.forEach((a) => {
    if (a.amount > 0 && Number.isFinite(a.amount)) {
      const z = Math.abs(a.amount - med) / (spread || 1);
      if (z >= 4) {
        a.score += weights.outlier;
        a.reasons.push('amount outlier');
      } else if (z >= 2.5) {
        a.score += Math.round(weights.outlier / 2);
        a.reasons.push('amount deviation');
      }
    }
    if (a.invoiceNo && invoiceMap.get(a.invoiceNo)?.count > 1) {
      a.score += weights.duplicate;
      a.reasons.push('duplicate invoice');
    }
    if (picked.qty && picked.unitPrice && picked.lineTotal) {
      const qty = Number(String(a.row[picked.qty]).replace(/,/g, ''));
      const unit = Number(String(a.row[picked.unitPrice]).replace(/,/g, ''));
      const total = Number(String(a.row[picked.lineTotal]).replace(/,/g, ''));
      if (Number.isFinite(qty) && Number.isFinite(unit) && Number.isFinite(total) && Math.abs(qty * unit - total) > Math.max(1, total * 0.02)) {
        a.score += weights.mismatch;
        a.reasons.push('qty*unit mismatch');
      }
    }
  });

  return {
    analyses: analyses.sort((a, b) => b.score - a.score).slice(0, 40),
    med,
    spread,
    totalCount: rows.length,
  };
}

export function aggStoreTod(headers: string[], rows: Row[], mapping: Mapping) {
  const picked = {
    invoiceNo: chooseHeader(headers, mapping.invoiceNo),
    invoiceAmount: chooseHeader(headers, mapping.invoiceAmount),
    store: chooseHeader(headers, mapping.store),
    date: chooseHeader(headers, mapping.date),
    time: chooseHeader(headers, mapping.time),
    lineTotal: chooseHeader(headers, mapping.lineTotal),
  };

  const storeMap = new Map<string, { count: number; amount: number; suspicious: number }>();
  const todMap = new Map<string, { count: number; amount: number }>();

  rows.forEach((row) => {
    const store = safeText(picked.store ? row[picked.store] : '') || 'Unknown';
    const rawDate = safeText(picked.date ? row[picked.date] : '');
    const rawTime = safeText(picked.time ? row[picked.time] : '');
    const amountRaw = picked.invoiceAmount ? row[picked.invoiceAmount] : (picked.lineTotal ? row[picked.lineTotal] : '');
    const amount = typeof amountRaw === 'number' ? amountRaw : Number(String(amountRaw).replace(/,/g, ''));
    const hour = parseDateParts(rawTime || rawDate).hour;
    const bucket = bucketHour(hour);
    const s = storeMap.get(store) ?? { count: 0, amount: 0, suspicious: 0 };
    s.count += 1;
    if (Number.isFinite(amount)) s.amount += amount;
    if (hour !== null && (hour < 6 || hour >= 22)) s.suspicious += 1;
    storeMap.set(store, s);
    const t = todMap.get(bucket) ?? { count: 0, amount: 0 };
    t.count += 1;
    if (Number.isFinite(amount)) t.amount += amount;
    todMap.set(bucket, t);
  });

  return {
    storeSummary: Array.from(storeMap.entries()).map(([store, v]) => ({ store, ...v })).sort((a, b) => b.amount - a.amount).slice(0, 12),
    todSummary: Array.from(todMap.entries()).map(([bucket, v]) => ({ bucket, ...v })),
  };
}

export function crossSheetRelations(sheets: { name: string; headers: string[]; rows: Row[] }[]): Relation[] {
  const relations: Relation[] = [];
  for (let i = 0; i < sheets.length; i += 1) {
    for (let j = i + 1; j < sheets.length; j += 1) {
      const left = sheets[i];
      const right = sheets[j];
      const leftValues = new Set(left.rows.flatMap((r) => Object.values(r).map((v) => safeText(v))).filter(Boolean));
      const rightValues = new Set(right.rows.flatMap((r) => Object.values(r).map((v) => safeText(v))).filter(Boolean));
      let overlap = 0;
      leftValues.forEach((v) => {
        if (rightValues.has(v)) overlap += 1;
      });
      if (overlap > 0) relations.push({ leftSheet: left.name, rightSheet: right.name, key: 'shared values', overlap, leftOnly: leftValues.size, rightOnly: rightValues.size });
    }
  }
  return relations.sort((a, b) => b.overlap - a.overlap).slice(0, 18);
}

export function buildReport(workbook: { sourceName: string; parsedAt: string; sheets: { name: string; rows: Row[]; anomalies: RowAnalysis[] }[]; rowCount: number; anomalyCount: number; relations: Relation[]; }) {
  const lines: string[] = [];
  lines.push(`# DOIT Report`);
  lines.push(`Source: ${workbook.sourceName}`);
  lines.push(`Generated: ${workbook.parsedAt}`);
  lines.push(`Sheets: ${workbook.sheets.length}`);
  lines.push(`Rows: ${workbook.rowCount}`);
  lines.push(`Anomalies: ${workbook.anomalyCount}`);
  lines.push('');
  lines.push('## Sheet summary');
  workbook.sheets.forEach((s) => lines.push(`- ${s.name}: ${s.rows.length} rows, ${s.anomalies.length} anomalies`));
  lines.push('');
  lines.push('## Suspicious invoices');
  workbook.sheets.flatMap((s) => s.anomalies.slice(0, 5)).forEach((a) => lines.push(`- ${a.sheet} :: ${a.invoiceNo || 'N/A'} :: score ${a.score} :: ${a.reasons.join(', ')}`));
  lines.push('');
  lines.push('## Relations');
  workbook.relations.forEach((r) => lines.push(`- ${r.leftSheet} ↔ ${r.rightSheet} [${r.key}] overlap=${r.overlap}`));
  return lines.join('\n');
}
