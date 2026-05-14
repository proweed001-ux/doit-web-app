import JSZip from 'jszip';
import { useEffect, useMemo, useState, type CSSProperties } from 'react';

type Cell = string | number | boolean | null;
type Row = Record<string, Cell>;
type Profile = 'balanced' | 'strict' | 'forensic';
type View = 'dashboard' | 'sheets' | 'relations' | 'tod' | 'reports' | 'database' | 'settings';
type CanonicalField = 'invoiceNo' | 'invoiceAmount' | 'store' | 'date' | 'time' | 'sku' | 'item' | 'qty' | 'unitPrice' | 'lineTotal' | 'customer' | 'category';

type Mapping = Partial<Record<CanonicalField, string>>;

type RowAnalysis = {
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

type SheetData = {
  name: string;
  headers: string[];
  rows: Row[];
  mapping: Mapping;
  anomalies: RowAnalysis[];
  numericSummary: { name: string; total: number; count: number }[];
  storeSummary: { store: string; count: number; amount: number; suspicious: number }[];
  todSummary: { bucket: string; count: number; amount: number }[];
};

type Relation = { leftSheet: string; rightSheet: string; key: string; overlap: number; leftOnly: number; rightOnly: number };
type HistoryItem = { name: string; parsedAt: string; sheets: number; rows: number; anomalies: number };
type Session = { name: string; org: string; token: string; createdAt: string };
type SavedReport = { title: string; generatedAt: string; markdown: string; source: string };
type DB = { sessions: Session[]; workbooks: HistoryItem[]; reports: SavedReport[]; templates: Record<string, Mapping> };
type Workbook = { sourceName: string; parsedAt: string; sheets: SheetData[]; warningCount: number; warnings: string[]; rowCount: number; anomalyCount: number; relations: Relation[]; reportMarkdown: string };

type Theme = { shell: string; panel: string; card: string; border: string; text: string; muted: string; accent: string; danger: string; success: string; tableAlt: string };

const AUTH_KEY = 'doit.auth.v2';
const DB_KEY = 'doit.db.v2';
const SETTINGS_KEY = 'doit.settings.v2';

const FIELD_RULES: Record<CanonicalField, RegExp[]> = {
  invoiceNo: [/invoice.*no/i, /invoice_no/i, /invoiceno/i, /inv\.?\s*no/i, /bill.*no/i, /doc.*no/i],
  invoiceAmount: [/invoice.*amt/i, /\bamount\b/i, /\bamt\b/i, /\btotal\b/i, /ยอดรวม/i, /value/i],
  store: [/store/i, /branch/i, /shop/i, /outlet/i, /location/i, /site/i],
  date: [/date/i, /doc.*date/i, /transaction.*date/i, /posting.*date/i, /วัน/i],
  time: [/time/i, /hour/i, /hh:?mm/i, /เวลา/i],
  sku: [/sku/i, /product.*code/i, /item.*code/i, /code/i, /barcode/i],
  item: [/item/i, /product/i, /description/i, /desc/i, /detail/i],
  qty: [/qty/i, /quantity/i, /pcs/i, /units?/i, /จำนวน/i],
  unitPrice: [/unit.*price/i, /price/i, /cost/i, /rate/i],
  lineTotal: [/line.*total/i, /net.*amount/i, /extended/i, /subtotal/i, /ยอดรวมบรรทัด/i],
  customer: [/customer/i, /client/i, /buyer/i, /account/i, /cust/i],
  category: [/category/i, /group/i, /type/i, /class/i],
};

const scoringProfiles: Record<Profile, Record<string, number>> = {
  balanced: { missing: 5, duplicate: 24, mismatch: 18, outlier: 16, relation: 10, oddHour: 4, blank: 3 },
  strict: { missing: 8, duplicate: 30, mismatch: 24, outlier: 22, relation: 14, oddHour: 6, blank: 5 },
  forensic: { missing: 10, duplicate: 36, mismatch: 28, outlier: 28, relation: 18, oddHour: 8, blank: 7 },
};

function loadJson<T>(key: string, fallback: T): T {
  if (typeof window === 'undefined') return fallback;
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback;
  }
}

function saveJson(key: string, value: unknown) {
  if (typeof window === 'undefined') return;
  window.localStorage.setItem(key, JSON.stringify(value));
}

function safeText(value: unknown): string {
  return String(value ?? '').replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim();
}

function normalize(value: unknown): string {
  return safeText(value)
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[\s\-_\/().,|]+/g, '')
    .replace(/[^\p{L}\p{N}]+/gu, '');
}

function colIndex(ref: string): number {
  const match = /([A-Z]+)/i.exec(ref);
  if (!match) return 0;
  return match[1].toUpperCase().split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
}

function colName(index: number): string {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - mod) / 26);
  }
  return s;
}

function splitCsv(line: string): string[] {
  const out: string[] = [];
  let current = '';
  let quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (quoted && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        quoted = !quoted;
      }
      continue;
    }
    if (ch === ',' && !quoted) {
      out.push(current);
      current = '';
      continue;
    }
    current += ch;
  }
  out.push(current);
  return out.map((v) => v.trim());
}

function parseDateParts(text: string): { date: string; time: string; hour: number | null } {
  const raw = safeText(text);
  if (!raw) return { date: '', time: '', hour: null };
  const m = raw.match(/(\d{1,2})[:.](\d{2})(?::(\d{2}))?/);
  if (m) {
    const hour = Number(m[1]);
    return { date: raw, time: `${m[1].padStart(2, '0')}:${m[2]}`, hour: Number.isFinite(hour) ? hour : null };
  }
  const d = new Date(raw);
  if (!Number.isNaN(d.getTime())) {
    return {
      date: d.toISOString().slice(0, 10),
      time: d.toTimeString().slice(0, 8),
      hour: d.getHours(),
    };
  }
  return { date: raw, time: '', hour: null };
}

function bucketHour(hour: number | null): string {
  if (hour === null || Number.isNaN(hour)) return 'Unknown';
  if (hour < 6) return 'Night';
  if (hour < 12) return 'Morning';
  if (hour < 17) return 'Afternoon';
  if (hour < 21) return 'Evening';
  return 'Late Night';
}

function makeTemplate(headers: string[]): Mapping {
  const found = (field: CanonicalField): string | undefined => headers.find((h) => FIELD_RULES[field].some((r) => r.test(h)));
  const mapping: Mapping = {};
  (Object.keys(FIELD_RULES) as CanonicalField[]).forEach((field) => {
    mapping[field] = found(field);
  });
  return mapping;
}

function chooseHeader(headers: string[], value: string | undefined | null): string | undefined {
  if (!value) return undefined;
  return headers.find((h) => h === value) ?? undefined;
}

function parseSharedStrings(xml: string): string[] {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return Array.from(doc.getElementsByTagName('si')).map((el) => safeText(el.textContent ?? ''));
}

function parseWorkbookRefs(workbookXml: string, relsXml: string): { name: string; path: string }[] {
  const wb = new DOMParser().parseFromString(workbookXml, 'application/xml');
  const rels = new DOMParser().parseFromString(relsXml, 'application/xml');
  const map = new Map<string, string>();
  Array.from(rels.getElementsByTagName('Relationship')).forEach((rel) => {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (id && target) map.set(id, target.startsWith('/') ? target.slice(1) : target);
  });
  return Array.from(wb.getElementsByTagName('sheet')).map((sheet) => {
    const name = sheet.getAttribute('name') ?? 'Sheet';
    const rid = sheet.getAttribute('r:id') ?? '';
    const target = map.get(rid) ?? 'worksheets/sheet1.xml';
    return { name, path: target.startsWith('xl/') ? target : `xl/${target}` };
  });
}

function parseWorksheet(xml: string, sharedStrings: string[], name: string): { name: string; headers: string[]; rows: Row[] } {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const matrix: Cell[][] = [];
  for (const row of Array.from(doc.getElementsByTagName('row'))) {
    const ri = Number(row.getAttribute('r') ?? matrix.length + 1) - 1;
    if (!matrix[ri]) matrix[ri] = [];
    for (const cell of Array.from(row.getElementsByTagName('c'))) {
      const ref = cell.getAttribute('r') ?? '';
      const ci = ref ? colIndex(ref) : matrix[ri].length;
      const t = cell.getAttribute('t') ?? '';
      const v = cell.getElementsByTagName('v')[0];
      const inline = cell.getElementsByTagName('is')[0];
      let value: Cell = null;
      if (t === 's') value = sharedStrings[Number(safeText(v?.textContent))] ?? '';
      else if (t === 'inlineStr') value = safeText(inline?.textContent ?? '');
      else if (t === 'b') value = safeText(v?.textContent) === '1';
      else {
        const raw = safeText(v?.textContent);
        value = raw === '' ? '' : Number.isFinite(Number(raw)) ? Number(raw) : raw;
      }
      matrix[ri][ci] = value;
    }
  }
  const firstHeader = matrix.findIndex((r) => r?.some((v) => safeText(v).length > 0));
  const headers = (matrix[firstHeader] ?? []).map((v, i) => safeText(v) || `Column ${colName(i)}`);
  const rows = matrix.slice(firstHeader + 1).filter((r) => r?.some((v) => safeText(v).length > 0)).map((r) => {
    const obj: Row = {};
    headers.forEach((h, i) => { obj[h] = r?.[i] ?? ''; });
    return obj;
  });
  return { name, headers, rows };
}

function parseCsvSheet(text: string, name: string): { name: string; headers: string[]; rows: Row[] } {
  const lines = text.split(/\r?\n/).filter((line) => safeText(line).length > 0);
  const headers = splitCsv(lines[0] ?? '').map((h, i) => safeText(h) || `Column ${colName(i)}`);
  const rows = lines.slice(1).map((line) => {
    const cells = splitCsv(line);
    const row: Row = {};
    headers.forEach((h, i) => { row[h] = cells[i] ?? ''; });
    return row;
  });
  return { name, headers, rows };
}

function numericSummary(rows: Row[], headers: string[]): { name: string; total: number; count: number }[] {
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

function pickHeaders(headers: string[], mapping: Mapping) {
  const resolve = (field: CanonicalField) => chooseHeader(headers, mapping[field]);
  return {
    invoiceNo: resolve('invoiceNo'),
    invoiceAmount: resolve('invoiceAmount'),
    store: resolve('store'),
    date: resolve('date'),
    time: resolve('time'),
    sku: resolve('sku'),
    item: resolve('item'),
    qty: resolve('qty'),
    unitPrice: resolve('unitPrice'),
    lineTotal: resolve('lineTotal'),
    customer: resolve('customer'),
    category: resolve('category'),
  };
}

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

function buildAnalysis(sheetName: string, headers: string[], rows: Row[], mapping: Mapping, profile: Profile) {
  const picked = pickHeaders(headers, mapping);
  const weights = scoringProfiles[profile];
  const invoiceMap = new Map<string, { count: number; amounts: number[]; indices: number[] }>();
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
      const g = invoiceMap.get(invoiceNo) ?? { count: 0, amounts: [], indices: [] };
      g.count += 1; g.indices.push(idx); if (Number.isFinite(amount)) g.amounts.push(amount);
      invoiceMap.set(invoiceNo, g);
    }
    const reasons: string[] = [];
    let score = 0;
    (['invoiceNo', 'invoiceAmount', 'store', 'date'] as CanonicalField[]).forEach((f) => {
      if (!picked[f] || !safeText(row[picked[f] as string])) { score += weights.missing; reasons.push(`missing ${f}`); }
    });
    if (!invoiceNo && picked.invoiceNo) { score += weights.blank; reasons.push('blank invoice'); }
    if (hour !== null && (hour < 6 || hour >= 22)) { score += weights.oddHour; reasons.push('odd hour'); }
    if (!Number.isFinite(amount) || amount <= 0) { score += weights.blank; reasons.push('invalid amount'); }
    analyses.push({ rowIndex: idx, score, reasons, invoiceNo, store, date: parsedTime.date || rawDate, time: parsedTime.time || rawTime, amount: Number.isFinite(amount) ? amount : 0, sheet: sheetName, row });
  });

  const med = median(amounts);
  const spread = mad(amounts, med);
  analyses.forEach((a) => {
    if (a.amount > 0 && Number.isFinite(a.amount)) {
      const z = Math.abs(a.amount - med) / (spread || 1);
      if (z >= 4) { a.score += weights.outlier; a.reasons.push('amount outlier'); }
      else if (z >= 2.5) { a.score += Math.round(weights.outlier / 2); a.reasons.push('amount deviation'); }
    }
    if (a.invoiceNo && invoiceMap.get(a.invoiceNo)?.count > 1) { a.score += weights.duplicate; a.reasons.push('duplicate invoice'); }
    if (picked.qty && picked.unitPrice && picked.lineTotal) {
      const qty = Number(String(a.row[picked.qty]).replace(/,/g, ''));
      const unit = Number(String(a.row[picked.unitPrice]).replace(/,/g, ''));
      const total = Number(String(a.row[picked.lineTotal]).replace(/,/g, ''));
      if (Number.isFinite(qty) && Number.isFinite(unit) && Number.isFinite(total) && Math.abs(qty * unit - total) > Math.max(1, total * 0.02)) {
        a.score += weights.mismatch; a.reasons.push('qty*unit mismatch');
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

function aggStoreTod(sheetName: string, headers: string[], rows: Row[], mapping: Mapping) {
  const picked = pickHeaders(headers, mapping);
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

function crossSheetRelations(sheets: SheetData[]): Relation[] {
  const keyTypes: { key: keyof ReturnType<typeof pickHeaders>; label: string }[] = [
    { key: 'invoiceNo', label: 'invoiceNo' },
    { key: 'store', label: 'store' },
    { key: 'date', label: 'date' },
    { key: 'customer', label: 'customer' },
    { key: 'sku', label: 'sku' },
  ];
  const relations: Relation[] = [];
  for (let i = 0; i < sheets.length; i += 1) {
    for (let j = i + 1; j < sheets.length; j += 1) {
      const left = sheets[i];
      const right = sheets[j];
      const leftSet = new Set(left.rows.map((r) => safeText(r[left.headers.find((h) => FIELD_RULES.invoiceNo.some((rx) => rx.test(h))) ?? '']) || '').filter(Boolean));
      const rightSet = new Set(right.rows.map((r) => safeText(r[right.headers.find((h) => FIELD_RULES.invoiceNo.some((rx) => rx.test(h))) ?? '']) || '').filter(Boolean));
      const leftValues = new Set(left.rows.flatMap((r) => Object.values(r).map((v) => safeText(v))).filter(Boolean));
      const rightValues = new Set(right.rows.flatMap((r) => Object.values(r).map((v) => safeText(v))).filter(Boolean));
      let overlap = 0;
      leftValues.forEach((v) => { if (rightValues.has(v)) overlap += 1; });
      if (overlap > 0) relations.push({ leftSheet: left.name, rightSheet: right.name, key: 'shared values', overlap, leftOnly: leftValues.size, rightOnly: rightValues.size });
      const invoiceOverlap = Array.from(leftSet).filter((v) => rightSet.has(v)).length;
      if (invoiceOverlap > 0) relations.push({ leftSheet: left.name, rightSheet: right.name, key: 'invoiceNo', overlap: invoiceOverlap, leftOnly: leftSet.size, rightOnly: rightSet.size });
    }
  }
  return relations.sort((a, b) => b.overlap - a.overlap).slice(0, 18);
}

function buildReport(workbook: Workbook): string {
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

function download(filename: string, content: string, type = 'text/plain;charset=utf-8') {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

async function parseWorkbook(file: File, profile: Profile, db: DB, session: Session | null): Promise<Workbook> {
  const warnings: string[] = [];
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const workbookXml = await zip.file('xl/workbook.xml')?.async('text');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('text');
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('text');

  if (!workbookXml || !relsXml) throw new Error('ไฟล์ไม่ใช่ workbook มาตรฐาน');
  const sharedStrings = parseSharedStrings(sharedStringsXml ?? '');
  const refs = parseWorkbookRefs(workbookXml, relsXml);
  const sheets: SheetData[] = [];
  const savedTemplates = db.templates;

  for (const ref of refs) {
    const xml = await zip.file(ref.path)?.async('text');
    if (!xml) { warnings.push(`Missing sheet data: ${ref.name}`); continue; }
    const raw = parseWorksheet(xml, sharedStrings, ref.name);
    const template = savedTemplates[ref.name] ?? makeTemplate(raw.headers);
    const mapping = { ...makeTemplate(raw.headers), ...template };
    const analysis = buildAnalysis(ref.name, raw.headers, raw.rows, mapping, profile);
    const agg = aggStoreTod(ref.name, raw.headers, raw.rows, mapping);
    sheets.push({
      name: ref.name,
      headers: raw.headers,
      rows: raw.rows,
      mapping,
      anomalies: analysis.analyses,
      numericSummary: numericSummary(raw.rows, raw.headers),
      storeSummary: agg.storeSummary,
      todSummary: agg.todSummary,
    });
    db.templates[ref.name] = mapping;
  }

  const relations = crossSheetRelations(sheets);
  const anomalyCount = sheets.reduce((sum, s) => sum + s.anomalies.length, 0);
  const rowCount = sheets.reduce((sum, s) => sum + s.rows.length, 0);
  const workbook: Workbook = {
    sourceName: file.name,
    parsedAt: new Date().toISOString(),
    sheets,
    warningCount: warnings.length,
    warnings,
    rowCount,
    anomalyCount,
    relations,
    reportMarkdown: '',
  };
  workbook.reportMarkdown = buildReport(workbook);
  if (session) {
    db.workbooks.unshift({ name: file.name, parsedAt: workbook.parsedAt, sheets: workbook.sheets.length, rows: workbook.rowCount, anomalies: workbook.anomalyCount });
    db.workbooks = db.workbooks.slice(0, 20);
  }
  return workbook;
}

function theme(dark: boolean): Theme {
  return dark
    ? { shell: '#0f172a', panel: '#111827', card: '#0b1220', border: '#334155', text: '#e2e8f0', muted: '#94a3b8', accent: '#38bdf8', danger: '#fb7185', success: '#34d399', tableAlt: 'rgba(255,255,255,0.03)' }
    : { shell: '#f8fafc', panel: '#ffffff', card: '#f1f5f9', border: '#cbd5e1', text: '#0f172a', muted: '#475569', accent: '#0284c7', danger: '#e11d48', success: '#059669', tableAlt: 'rgba(15,23,42,0.03)' };
}

export default function DoitAppPro() {
  const auth = loadJson<Session | null>(AUTH_KEY, null);
  const initialDB = loadJson<DB>(DB_KEY, { sessions: [], workbooks: [], reports: [], templates: {} });
  const settings = loadJson<{ dark: boolean; rowsLimit: number; profile: Profile }>(SETTINGS_KEY, { dark: true, rowsLimit: 100, profile: 'balanced' });

  const [session, setSession] = useState<Session | null>(auth);
  const [db, setDB] = useState<DB>(initialDB);
  const [dark, setDark] = useState(settings.dark);
  const [rowsLimit, setRowsLimit] = useState(settings.rowsLimit);
  const [profile, setProfile] = useState<Profile>(settings.profile);
  const [view, setView] = useState<View>('dashboard');
  const [status, setStatus] = useState('พร้อมใช้งาน');
  const [error, setError] = useState('');
  const [query, setQuery] = useState('');
  const [dragging, setDragging] = useState(false);
  const [workbook, setWorkbook] = useState<Workbook | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);

  const t = theme(dark);
  const active = workbook?.sheets[activeSheet] ?? null;
  const filteredRows = useMemo(() => {
    if (!active) return [] as Row[];
    let rows = active.rows;
    if (query.trim()) {
      const q = normalize(query);
      rows = rows.filter((row) => Object.entries(row).some(([k, v]) => normalize(k).includes(q) || normalize(v).includes(q)));
    }
    return rows;
  }, [active, query]);
  const visibleRows = filteredRows.slice(0, rowsLimit);
  const allAnomalies = workbook?.sheets.flatMap((s) => s.anomalies) ?? [];
  const todAll = workbook?.sheets.flatMap((s) => s.todSummary.map((x) => ({ ...x, sheet: s.name }))) ?? [];
  const storesAll = workbook?.sheets.flatMap((s) => s.storeSummary.map((x) => ({ ...x, sheet: s.name }))) ?? [];
  const latestReport = workbook?.reportMarkdown ?? '';

  useEffect(() => {
    saveJson(AUTH_KEY, session);
    saveJson(DB_KEY, db);
    saveJson(SETTINGS_KEY, { dark, rowsLimit, profile });
    if (typeof document !== 'undefined') {
      document.documentElement.style.colorScheme = dark ? 'dark' : 'light';
      document.body.style.margin = '0';
      document.body.style.background = t.shell;
    }
  }, [session, db, dark, rowsLimit, profile, t.shell]);

  function createSession(name: string, org: string) {
    const next: Session = { name: safeText(name) || 'User', org: safeText(org) || 'Workspace', token: crypto.randomUUID(), createdAt: new Date().toISOString() };
    setSession(next);
    setDB((prev) => ({ ...prev, sessions: [next, ...prev.sessions].slice(0, 20) }));
    setStatus(`เข้าสู่ระบบแล้ว: ${next.name}`);
  }

  function clearSession() {
    setSession(null);
    setStatus('ออกจากระบบ');
  }

  function saveTemplate() {
    if (!active || !workbook) return;
    setDB((prev) => ({ ...prev, templates: { ...prev.templates, [active.name]: active.mapping } }));
    setStatus(`บันทึก mapping ของ ${active.name}`);
  }

  async function onFile(file: File) {
    if (!session) {
      setError('กรุณาเข้าสู่ระบบก่อน');
      return;
    }
    setError('');
    setStatus(`กำลังอ่าน ${file.name}`);
    try {
      const wb = /\.(xlsx|xlsm|zip)$/i.test(file.name) ? await parseWorkbook(file, profile, { ...db }, session) : await (async () => {
        const csv = parseCsvSheet(await file.text(), file.name);
        const mapping = makeTemplate(csv.headers);
        const analysis = buildAnalysis(csv.name, csv.headers, csv.rows, mapping, profile);
        const agg = aggStoreTod(csv.name, csv.headers, csv.rows, mapping);
        return {
          sourceName: file.name,
          parsedAt: new Date().toISOString(),
          sheets: [{ name: csv.name, headers: csv.headers, rows: csv.rows, mapping, anomalies: analysis.analyses, numericSummary: numericSummary(csv.rows, csv.headers), storeSummary: agg.storeSummary, todSummary: agg.todSummary }],
          warningCount: 1,
          warnings: ['CSV fallback mode'],
          rowCount: csv.rows.length,
          anomalyCount: analysis.analyses.length,
          relations: [],
          reportMarkdown: '',
        } as Workbook;
      })();
      setWorkbook(wb);
      setActiveSheet(0);
      setView('dashboard');
      setStatus(`โหลดสำเร็จ: ${wb.sheets.length} sheet`);
      const report: SavedReport = { title: wb.sourceName, generatedAt: wb.parsedAt, markdown: wb.reportMarkdown, source: wb.sourceName };
      setDB((prev) => ({ ...prev, reports: [report, ...prev.reports].slice(0, 20), workbooks: [{ name: wb.sourceName, parsedAt: wb.parsedAt, sheets: wb.sheets.length, rows: wb.rowCount, anomalies: wb.anomalyCount }, ...prev.workbooks].slice(0, 20) }));
    } catch (e) {
      setError((e as Error).message || 'ไม่สามารถอ่านไฟล์ได้');
      setStatus('เกิดข้อผิดพลาด');
      setWorkbook(null);
    }
  }

  const btn = (id: View) => ({
    padding: '10px 14px',
    borderRadius: 999,
    border: `1px solid ${view === id ? t.accent : t.border}`,
    background: view === id ? `${t.accent}22` : t.panel,
    color: t.text,
    cursor: 'pointer',
    fontWeight: 700 as const,
  });

  const shell: CSSProperties = { minHeight: '100vh', background: t.shell, color: t.text, fontFamily: 'Inter, system-ui, sans-serif' };
  const panel: CSSProperties = { background: t.panel, border: `1px solid ${t.border}`, borderRadius: 20, boxShadow: dark ? '0 20px 60px rgba(0,0,0,0.18)' : '0 18px 48px rgba(15,23,42,0.06)' };
  const card: CSSProperties = { background: t.card, border: `1px solid ${t.border}`, borderRadius: 16 };

  const fileSummary = workbook ? [
    { label: 'Sheets', value: workbook.sheets.length },
    { label: 'Rows', value: workbook.rowCount },
    { label: 'Anomalies', value: workbook.anomalyCount },
    { label: 'Relations', value: workbook.relations.length },
  ] : [];

  function exportReport() {
    if (!workbook) return;
    download(`${workbook.sourceName.replace(/\.[^.]+$/, '')}.md`, workbook.reportMarkdown, 'text/markdown;charset=utf-8');
  }

  function exportWorkbookJson() {
    if (!workbook) return;
    download(`${workbook.sourceName.replace(/\.[^.]+$/, '')}.json`, JSON.stringify(workbook, null, 2), 'application/json;charset=utf-8');
  }

  function updateMapping(field: CanonicalField, header: string) {
    if (!active) return;
    const mapping = { ...active.mapping, [field]: header || undefined };
    const next = { ...workbook! };
    next.sheets = next.sheets.map((s, idx) => idx === activeSheet ? { ...s, mapping, ...recomputeSheet(s.name, s.headers, s.rows, mapping, profile) } : s);
    next.relations = crossSheetRelations(next.sheets);
    next.anomalyCount = next.sheets.reduce((sum, s) => sum + s.anomalies.length, 0);
    next.rowCount = next.sheets.reduce((sum, s) => sum + s.rows.length, 0);
    next.reportMarkdown = buildReport(next);
    setWorkbook(next);
    setDB((prev) => ({ ...prev, templates: { ...prev.templates, [active.name]: mapping } }));
  }

  function recomputeSheet(name: string, headers: string[], rows: Row[], mapping: Mapping, profileNow: Profile) {
    const analysis = buildAnalysis(name, headers, rows, mapping, profileNow);
    const agg = aggStoreTod(name, headers, rows, mapping);
    return { anomalies: analysis.analyses, numericSummary: numericSummary(rows, headers), storeSummary: agg.storeSummary, todSummary: agg.todSummary };
  }

  return (
    <div style={shell}>
      <div style={{ maxWidth: 1480, margin: '0 auto', padding: 20 }}>
        {!session ? (
          <div style={{ ...panel, maxWidth: 760, margin: '48px auto', padding: 24 }}>
            <h1 style={{ margin: 0, fontSize: 'clamp(2rem, 4vw, 3rem)' }}>DOIT Workspace</h1>
            <p style={{ color: t.muted }}>เข้าระบบเพื่อใช้ engine วิเคราะห์ workbook, dashboard, TOD/store aggregation, relations และ report generator</p>
            <div style={{ display: 'grid', gap: 12, marginTop: 18 }}>
              <input id="name" placeholder="ชื่อผู้ใช้" style={{ padding: 12, borderRadius: 12, border: `1px solid ${t.border}`, background: t.card, color: t.text }} />
              <input id="org" placeholder="ทีม/องค์กร" style={{ padding: 12, borderRadius: 12, border: `1px solid ${t.border}`, background: t.card, color: t.text }} />
              <button onClick={() => createSession((document.getElementById('name') as HTMLInputElement)?.value ?? '', (document.getElementById('org') as HTMLInputElement)?.value ?? '')} style={{ padding: 12, borderRadius: 12, border: 'none', background: t.accent, color: '#082f49', fontWeight: 800, cursor: 'pointer' }}>เข้าสู่ระบบ</button>
            </div>
          </div>
        ) : (
          <>
            <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap', alignItems: 'flex-end' }}>
              <div>
                <div style={{ fontSize: 13, color: t.muted, letterSpacing: 1.2, textTransform: 'uppercase' }}>DOIT Web App</div>
                <h1 style={{ margin: '8px 0 6px', fontSize: 'clamp(2rem, 3vw, 3.4rem)', lineHeight: 1.05 }}>Workbook Control Center</h1>
                <div style={{ color: t.muted }}>Signed in as <strong>{session.name}</strong> · {session.org}</div>
              </div>
              <div style={{ ...panel, minWidth: 280, padding: 16 }}>
                <div style={{ fontSize: 12, color: t.muted }}>Status</div>
                <div style={{ fontWeight: 800 }}>{status}</div>
                {workbook && <div style={{ marginTop: 6, fontSize: 12, color: t.muted }}>{workbook.sourceName} · {new Date(workbook.parsedAt).toLocaleString()}</div>}
              </div>
            </div>

            <div style={{ marginTop: 18, ...panel, padding: 18 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', gap: 16, flexWrap: 'wrap', alignItems: 'center' }}>
                <div>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>Import</div>
                  <div style={{ color: t.muted }}>รองรับ .xlsx, .xlsm, .zip, .csv</div>
                </div>
                <label style={{ cursor: 'pointer', padding: '12px 16px', borderRadius: 14, background: t.accent, color: '#082f49', fontWeight: 800 }}>
                  Choose file
                  <input type="file" accept=".xlsx,.xlsm,.zip,.csv,.txt" style={{ display: 'none' }} onChange={(e) => { const f = e.target.files?.[0]; if (f) void onFile(f); }} />
                </label>
              </div>
              {dragging && <div style={{ marginTop: 12, color: t.accent }}>Drop file to import</div>}
              {error && <div style={{ marginTop: 12, padding: 12, borderRadius: 12, background: `${t.danger}18`, color: t.danger }}>{error}</div>}
            </div>

            <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 16 }}>
              <button style={btn('dashboard')} onClick={() => setView('dashboard')}>Dashboard</button>
              <button style={btn('sheets')} onClick={() => setView('sheets')}>Sheets</button>
              <button style={btn('relations')} onClick={() => setView('relations')}>Relations</button>
              <button style={btn('tod')} onClick={() => setView('tod')}>TOD / Store</button>
              <button style={btn('reports')} onClick={() => setView('reports')}>Reports</button>
              <button style={btn('database')} onClick={() => setView('database')}>Database</button>
              <button style={btn('settings')} onClick={() => setView('settings')}>Settings</button>
              <button style={btn('dashboard')} onClick={exportReport} disabled={!workbook}>Export Report</button>
              <button style={btn('dashboard')} onClick={exportWorkbookJson} disabled={!workbook}>Export JSON</button>
              <button style={btn('dashboard')} onClick={clearSession}>Logout</button>
            </div>

            {workbook && view === 'dashboard' && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 16 }}>
                  {fileSummary.map((s) => (
                    <div key={s.label} style={{ ...card, padding: 18 }}><div style={{ color: t.muted, fontSize: 12 }}>{s.label}</div><div style={{ fontSize: 28, fontWeight: 900, marginTop: 4 }}>{s.value}</div></div>
                  ))}
                </div>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>Top anomalies</div>
                  <div style={{ marginTop: 12, display: 'grid', gap: 10 }}>
                    {allAnomalies.slice(0, 12).map((a) => (
                      <div key={`${a.sheet}-${a.rowIndex}`} style={{ padding: 12, borderRadius: 12, background: card.background, border: `1px solid ${t.border}` }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                          <strong>{a.sheet} · {a.invoiceNo || `Row ${a.rowIndex + 1}`}</strong>
                          <span style={{ color: t.danger, fontWeight: 800 }}>score {a.score}</span>
                        </div>
                        <div style={{ color: t.muted, marginTop: 4, fontSize: 12 }}>{a.reasons.join(' · ') || 'normal'}</div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {workbook && view === 'sheets' && active && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                  {workbook.sheets.map((sheet, idx) => <button key={sheet.name + idx} onClick={() => setActiveSheet(idx)} style={{ ...btn('sheets'), background: idx === activeSheet ? `${t.accent}22` : t.panel }}>{sheet.name}</button>)}
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1.6fr) minmax(280px, 0.9fr)', gap: 18 }}>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ ...panel, padding: 18 }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                        <div><div style={{ fontSize: 12, color: t.muted }}>Active sheet</div><div style={{ fontSize: 20, fontWeight: 800 }}>{active.name}</div></div>
                        <div style={{ color: t.muted }}>{active.rows.length} rows · {active.headers.length} columns</div>
                      </div>
                      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 14 }}>
                        <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search..." style={{ flex: '1 1 220px', padding: 12, borderRadius: 12, border: `1px solid ${t.border}`, background: t.card, color: t.text }} />
                        <button onClick={() => saveTemplate()} style={{ ...btn('sheets') }}>Save mapping</button>
                      </div>
                    </div>
                    <div style={{ marginTop: 16, overflowX: 'auto', borderRadius: 18, border: `1px solid ${t.border}`, background: t.panel }}>
                      <table style={{ width: '100%', minWidth: 900 }}>
                        <thead>
                          <tr>{active.headers.map((h) => <th key={h} style={{ position: 'sticky', top: 0, background: t.panel, textAlign: 'left', padding: 12, borderBottom: `1px solid ${t.border}`, color: t.muted }}>{h}</th>)}</tr>
                        </thead>
                        <tbody>
                          {visibleRows.map((row, i) => (
                            <tr key={i} style={{ background: i % 2 ? t.tableAlt : 'transparent' }}>{active.headers.map((h) => <td key={h} style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{safeText(row[h]) || '—'}</td>)}</tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                  <div style={{ display: 'grid', gap: 18 }}>
                    <div style={{ ...panel, padding: 18 }}>
                      <div style={{ fontWeight: 800, fontSize: 18 }}>Auto mapping template</div>
                      <div style={{ display: 'grid', gap: 10, marginTop: 12 }}>
                        {(Object.keys(FIELD_RULES) as CanonicalField[]).map((field) => (
                          <label key={field} style={{ display: 'grid', gap: 6 }}>
                            <span style={{ color: t.muted, fontSize: 12 }}>{field}</span>
                            <select value={active.mapping[field] ?? ''} onChange={(e) => updateMapping(field, e.target.value)} style={{ padding: 12, borderRadius: 12, border: `1px solid ${t.border}`, background: t.card, color: t.text }}>
                              <option value="">(none)</option>
                              {active.headers.map((h) => <option key={h} value={h}>{h}</option>)}
                            </select>
                          </label>
                        ))}
                      </div>
                    </div>
                    <div style={{ ...panel, padding: 18 }}>
                      <div style={{ fontWeight: 800, fontSize: 18 }}>Sheet summary</div>
                      <div style={{ display: 'grid', gap: 8, marginTop: 12 }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Rows</span><strong>{active.rows.length}</strong></div>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Anomalies</span><strong>{active.anomalies.length}</strong></div>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Stores</span><strong>{active.storeSummary.length}</strong></div>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>TOD buckets</span><strong>{active.todSummary.length}</strong></div>
                      </div>
                    </div>
                    <div style={{ ...panel, padding: 18 }}>
                      <div style={{ fontWeight: 800, fontSize: 18 }}>Top numeric columns</div>
                      <div style={{ display: 'grid', gap: 10, marginTop: 12 }}>
                        {active.numericSummary.slice(0, 8).map((n) => <div key={n.name} style={{ padding: 12, borderRadius: 12, background: t.card, border: `1px solid ${t.border}` }}><div style={{ color: t.muted, fontSize: 12 }}>{n.name}</div><strong>{Number.isInteger(n.total) ? n.total.toLocaleString() : n.total.toFixed(2)}</strong><div style={{ color: t.muted, fontSize: 12 }}>{n.count} numeric cells</div></div>)}
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {workbook && view === 'relations' && (
              <div style={{ marginTop: 22, ...panel, padding: 18 }}>
                <div style={{ fontWeight: 800, fontSize: 18 }}>Cross-sheet relation engine</div>
                <div style={{ marginTop: 12, display: 'grid', gap: 10 }}>
                  {workbook.relations.length ? workbook.relations.map((r, idx) => <div key={idx} style={{ padding: 12, borderRadius: 12, background: card.background, border: `1px solid ${t.border}` }}><strong>{r.leftSheet} ↔ {r.rightSheet}</strong><div style={{ color: t.muted, fontSize: 12 }}>key: {r.key} · overlap: {r.overlap} · left: {r.leftOnly} · right: {r.rightOnly}</div></div>) : <div style={{ color: t.muted }}>No strong cross-sheet relation detected</div>}
                </div>
              </div>
            )}

            {workbook && view === 'tod' && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>TOD / Store aggregation</div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 16, marginTop: 12 }}>
                    {storesAll.slice(0, 8).map((s, idx) => <div key={idx} style={{ ...card, padding: 14 }}><div style={{ fontWeight: 800 }}>{s.store}</div><div style={{ color: t.muted, fontSize: 12 }}>{s.sheet}</div><div style={{ marginTop: 8 }}>count {s.count}</div><div>amount {Number.isInteger(s.amount) ? s.amount.toLocaleString() : s.amount.toFixed(2)}</div><div style={{ color: t.danger }}>suspicious {s.suspicious}</div></div>)}
                  </div>
                </div>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>TOD buckets</div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12, marginTop: 12 }}>
                    {todAll.map((x, idx) => <div key={idx} style={{ ...card, padding: 14 }}><div style={{ fontWeight: 800 }}>{x.bucket}</div><div style={{ color: t.muted, fontSize: 12 }}>{x.sheet}</div><div>count {x.count}</div><div>amount {Number.isInteger(x.amount) ? x.amount.toLocaleString() : x.amount.toFixed(2)}</div></div>)}
                  </div>
                </div>
              </div>
            )}

            {workbook && view === 'reports' && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                    <div>
                      <div style={{ fontWeight: 800, fontSize: 18 }}>Report generator</div>
                      <div style={{ color: t.muted, marginTop: 4 }}>Markdown summary of workbook, anomalies and relations</div>
                    </div>
                    <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                      <button onClick={exportReport} style={{ ...btn('reports') }}>Download MD</button>
                      <button onClick={exportWorkbookJson} style={{ ...btn('reports') }}>Download JSON</button>
                    </div>
                  </div>
                  <pre style={{ whiteSpace: 'pre-wrap', marginTop: 16, padding: 16, borderRadius: 14, background: t.card, border: `1px solid ${t.border}`, overflowX: 'auto' }}>{latestReport}</pre>
                </div>
              </div>
            )}

            {view === 'database' && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>Persistent local database</div>
                  <div style={{ color: t.muted, marginTop: 4 }}>Sessions, workbooks, reports and templates are saved in browser storage and can be exported/imported.</div>
                  <div style={{ display: 'grid', gap: 10, marginTop: 12 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Sessions</span><strong>{db.sessions.length}</strong></div>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Workbooks</span><strong>{db.workbooks.length}</strong></div>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Reports</span><strong>{db.reports.length}</strong></div>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Templates</span><strong>{Object.keys(db.templates).length}</strong></div>
                  </div>
                  <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 14 }}>
                    <button onClick={() => download('doit-database.json', JSON.stringify(db, null, 2), 'application/json;charset=utf-8')} style={{ ...btn('database') }}>Export DB</button>
                    <button onClick={() => { setDB({ sessions: [], workbooks: [], reports: [], templates: {} }); setStatus('Database cleared'); }} style={{ ...btn('database') }}>Clear DB</button>
                  </div>
                </div>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>Session</div>
                  <div style={{ marginTop: 10 }}>User: <strong>{session.name}</strong></div>
                  <div>Org: <strong>{session.org}</strong></div>
                  <div>Token: <code>{session.token.slice(0, 8)}...</code></div>
                  <div style={{ marginTop: 12, display: 'flex', gap: 10, flexWrap: 'wrap' }}><button onClick={clearSession} style={{ ...btn('database') }}>Logout</button></div>
                </div>
              </div>
            )}

            {view === 'settings' && (
              <div style={{ marginTop: 22, ...panel, padding: 18 }}>
                <div style={{ fontWeight: 800, fontSize: 18 }}>Settings & scoring profiles</div>
                <div style={{ display: 'grid', gap: 16, marginTop: 14, maxWidth: 760 }}>
                  <label style={{ display: 'grid', gap: 8, padding: 14, borderRadius: 14, background: t.card, border: `1px solid ${t.border}` }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>Dark mode</span><strong>{String(dark)}</strong></div>
                    <input type="checkbox" checked={dark} onChange={(e) => setDark(e.target.checked)} />
                  </label>
                  <label style={{ display: 'grid', gap: 8, padding: 14, borderRadius: 14, background: t.card, border: `1px solid ${t.border}` }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>Rows shown</span><strong>{rowsLimit}</strong></div>
                    <input type="range" min={25} max={500} step={25} value={rowsLimit} onChange={(e) => setRowsLimit(Number(e.target.value))} />
                  </label>
                  <label style={{ display: 'grid', gap: 8, padding: 14, borderRadius: 14, background: t.card, border: `1px solid ${t.border}` }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>Scoring profile</span><strong>{profile}</strong></div>
                    <select value={profile} onChange={(e) => setProfile(e.target.value as Profile)} style={{ padding: 12, borderRadius: 12, border: `1px solid ${t.border}`, background: t.panel, color: t.text }}>
                      <option value="balanced">balanced</option>
                      <option value="strict">strict</option>
                      <option value="forensic">forensic</option>
                    </select>
                  </label>
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
