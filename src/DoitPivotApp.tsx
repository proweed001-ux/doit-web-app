import JSZip from 'jszip';
import { useMemo, useState } from 'react';

type Cell = string | number | boolean | null;
type RawRow = Record<string, Cell>;
type FieldKey = 'sellerGroup' | 'brand' | 'tasSizeGroup' | 'sku' | 'description' | 'soTypeId' | 'shipQtyPCS' | 'invoiceAmt' | 'date';
type FieldMap = Partial<Record<FieldKey, string>>;
type ParsedSheet = { name: string; headers: string[]; rows: RawRow[] };
type Workbook = { name: string; parsedAt: string; sheets: ParsedSheet[]; warnings: string[] };
type InvoiceLine = {
  sourceRowNo: number;
  sellerGroup: string;
  brand: string;
  tasSizeGroup: string;
  sku: string;
  description: string;
  soTypeId: string;
  date: string;
  shipQtyPCS: number;
  invoiceAmt: number;
  raw: RawRow;
};
type PivotRow = Omit<InvoiceLine, 'sourceRowNo' | 'date' | 'raw'> & { lineCount: number };

type Theme = {
  shell: string;
  panel: string;
  card: string;
  text: string;
  muted: string;
  border: string;
  accent: string;
  danger: string;
  good: string;
  tableAlt: string;
};

const FIELD_LABELS: Record<FieldKey, string> = {
  sellerGroup: 'SO Seller Group / PS',
  brand: 'Brand',
  tasSizeGroup: 'TAS_SizeGroup',
  sku: 'SKU',
  description: 'Description',
  soTypeId: 'SOTypeID',
  shipQtyPCS: 'ShipQtyPCS',
  invoiceAmt: 'InvoiceAmt',
  date: 'SO_Date / Invoice_Date',
};

const FIELD_ALIASES: Record<FieldKey, string[]> = {
  sellerGroup: ['sosellergroup', 'sellergroup', 'ps', 'psname', 'salesperson', 'salesgroup'],
  brand: ['brand', 'brandname'],
  tasSizeGroup: ['tassizegroup', 'tasizegroup', 'sizegroup', 'tasgroup'],
  sku: ['sku', 'skucode', 'itemcode', 'productcode', 'material', 'materialcode'],
  description: ['description', 'desc', 'skudescription', 'productdescription', 'itemdescription'],
  soTypeId: ['sotypeid', 'sotype', 'sotypecode', 'doctype', 'documenttype'],
  shipQtyPCS: ['shipqtypcs', 'shipqtypc', 'shipqtypcs', 'shipqtypc', 'shipqty', 'qtypcs'],
  invoiceAmt: ['invoiceamt', 'invoiceamount'],
  date: ['sodate', 'invoicedate', 'docdate', 'documentdate', 'date'],
};

const GROUP_FIELDS: FieldKey[] = ['sellerGroup', 'brand', 'tasSizeGroup', 'sku', 'description', 'soTypeId'];
const REQUIRED_FIELDS: FieldKey[] = ['sellerGroup', 'brand', 'tasSizeGroup', 'sku', 'description', 'soTypeId', 'shipQtyPCS', 'invoiceAmt'];
const DISALLOWED_AMOUNT_HEADERS = ['totinvc', 'totalinvoice', 'invoicetotal', 'totalinvc', 'headeramount', 'billtotal'];
const MAX_FILE_SIZE_MB = 35;
const MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024;

function theme(): Theme {
  return {
    shell: '#f8fafc',
    panel: '#ffffff',
    card: '#f1f5f9',
    text: '#0f172a',
    muted: '#475569',
    border: '#cbd5e1',
    accent: '#0369a1',
    danger: '#b91c1c',
    good: '#047857',
    tableAlt: '#f8fafc',
  };
}

function safeText(value: unknown): string {
  return String(value ?? '').replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim();
}

function normalize(value: unknown): string {
  return safeText(value).toLowerCase().normalize('NFKD').replace(/[^\p{L}\p{N}]+/gu, '');
}

function formatNumber(value: number, digits = 2): string {
  if (!Number.isFinite(value)) return '0';
  return value.toLocaleString('th-TH', { maximumFractionDigits: digits });
}

function parseNumber(value: Cell): number {
  if (typeof value === 'number') return Number.isFinite(value) ? value : Number.NaN;
  const raw = safeText(value);
  if (!raw) return Number.NaN;
  const negativeByParen = /^\(.*\)$/.test(raw);
  const cleaned = raw.replace(/[(),\s]/g, '').replace(/−/g, '-');
  const n = Number(cleaned);
  if (!Number.isFinite(n)) return Number.NaN;
  return negativeByParen && n > 0 ? -n : n;
}

function excelSerialToDate(value: number): string {
  if (!Number.isFinite(value) || value < 20000 || value > 70000) return '';
  const utc = Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000;
  return new Date(utc).toISOString().slice(0, 10);
}

function dateText(value: Cell): string {
  if (typeof value === 'number') return excelSerialToDate(value) || safeText(value);
  const raw = safeText(value);
  if (!raw) return '';
  const direct = raw.match(/\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}/)?.[0] ?? raw;
  const parsed = new Date(direct);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return raw;
}

function csvEscape(value: unknown): string {
  const text = safeText(value);
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function downloadText(name: string, text: string, mime = 'text/plain;charset=utf-8') {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
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
      out.push(current.trim());
      current = '';
      continue;
    }
    current += ch;
  }
  out.push(current.trim());
  return out;
}

function colIndex(ref: string): number {
  const match = /([A-Z]+)/i.exec(ref);
  if (!match) return 0;
  return match[1].toUpperCase().split('').reduce((sum, ch) => sum * 26 + ch.charCodeAt(0) - 64, 0) - 1;
}

function colName(index: number): string {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    out = String.fromCharCode(65 + mod) + out;
    n = Math.floor((n - mod) / 26);
  }
  return out;
}

function uniqueHeaders(values: Cell[]): string[] {
  const used = new Map<string, number>();
  return values.map((value, index) => {
    const base = safeText(value) || `Column ${colName(index)}`;
    const count = used.get(base) ?? 0;
    used.set(base, count + 1);
    return count ? `${base} (${count + 1})` : base;
  });
}

function detectFields(headers: string[]): FieldMap {
  const byNorm = new Map<string, string>();
  headers.forEach((header) => byNorm.set(normalize(header), header));
  const result: FieldMap = {};
  (Object.keys(FIELD_ALIASES) as FieldKey[]).forEach((field) => {
    const hit = FIELD_ALIASES[field].map((alias) => byNorm.get(alias)).find(Boolean);
    if (hit) result[field] = hit;
  });
  return result;
}

function headerScore(row: Cell[]): number {
  const headers = row.map(safeText).filter(Boolean);
  if (!headers.length) return 0;
  const detected = detectFields(headers);
  return REQUIRED_FIELDS.reduce((sum, field) => sum + (detected[field] ? 1 : 0), 0);
}

function findHeaderRow(matrix: Cell[][]): number {
  let bestIndex = matrix.findIndex((row) => row?.some((cell) => safeText(cell)));
  let bestScore = -1;
  matrix.forEach((row, index) => {
    const score = headerScore(row ?? []);
    if (score > bestScore) {
      bestIndex = index;
      bestScore = score;
    }
  });
  return Math.max(0, bestIndex);
}

function parseSharedStrings(xml: string): string[] {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return Array.from(doc.getElementsByTagName('si')).map((el) => safeText(el.textContent ?? ''));
}

function parseWorkbookRefs(workbookXml: string, relsXml: string): { name: string; path: string }[] {
  const wb = new DOMParser().parseFromString(workbookXml, 'application/xml');
  const rels = new DOMParser().parseFromString(relsXml, 'application/xml');
  const relMap = new Map<string, string>();
  Array.from(rels.getElementsByTagName('Relationship')).forEach((rel) => {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (id && target) relMap.set(id, target.startsWith('/') ? target.slice(1) : target);
  });
  return Array.from(wb.getElementsByTagName('sheet')).map((sheet, index) => {
    const name = sheet.getAttribute('name') ?? `Sheet${index + 1}`;
    const relId = sheet.getAttribute('r:id') ?? '';
    const target = relMap.get(relId) ?? `worksheets/sheet${index + 1}.xml`;
    return { name, path: target.startsWith('xl/') ? target : `xl/${target}` };
  });
}

function parseWorksheet(xml: string, sharedStrings: string[], name: string): ParsedSheet {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const matrix: Cell[][] = [];
  Array.from(doc.getElementsByTagName('row')).forEach((row) => {
    const rowIndex = Number(row.getAttribute('r') ?? matrix.length + 1) - 1;
    if (!matrix[rowIndex]) matrix[rowIndex] = [];
    Array.from(row.getElementsByTagName('c')).forEach((cell) => {
      const ref = cell.getAttribute('r') ?? '';
      const cellIndex = ref ? colIndex(ref) : matrix[rowIndex].length;
      const type = cell.getAttribute('t') ?? '';
      const v = cell.getElementsByTagName('v')[0];
      const inline = cell.getElementsByTagName('is')[0];
      let value: Cell = '';
      if (type === 's') value = sharedStrings[Number(safeText(v?.textContent))] ?? '';
      else if (type === 'inlineStr') value = safeText(inline?.textContent ?? '');
      else if (type === 'b') value = safeText(v?.textContent) === '1';
      else {
        const raw = safeText(v?.textContent ?? cell.textContent ?? '');
        value = raw === '' ? '' : Number.isFinite(Number(raw)) ? Number(raw) : raw;
      }
      matrix[rowIndex][cellIndex] = value;
    });
  });
  const headerIndex = findHeaderRow(matrix);
  const headers = uniqueHeaders(matrix[headerIndex] ?? []);
  const rows = matrix.slice(headerIndex + 1).filter((row) => row?.some((cell) => safeText(cell))).map((row) => {
    const record: RawRow = {};
    headers.forEach((header, index) => {
      record[header] = row?.[index] ?? '';
    });
    return record;
  });
  return { name, headers, rows };
}

function parseCsvSheet(text: string, name: string): ParsedSheet {
  const lines = text.split(/\r?\n/).filter((line) => safeText(line));
  const rawMatrix = lines.map(splitCsv);
  const headerIndex = findHeaderRow(rawMatrix);
  const headers = uniqueHeaders(rawMatrix[headerIndex] ?? []);
  const rows = rawMatrix.slice(headerIndex + 1).filter((row) => row.some((cell) => safeText(cell))).map((cells) => {
    const record: RawRow = {};
    headers.forEach((header, index) => {
      record[header] = cells[index] ?? '';
    });
    return record;
  });
  return { name, headers, rows };
}

async function parseFile(file: File): Promise<Workbook> {
  if (/\.csv$/i.test(file.name) || /\.txt$/i.test(file.name)) {
    return { name: file.name, parsedAt: new Date().toISOString(), sheets: [parseCsvSheet(await file.text(), file.name)], warnings: ['อ่านเป็นไฟล์ CSV/TXT'] };
  }
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const workbookXml = await zip.file('xl/workbook.xml')?.async('text');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('text');
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('text');
  if (!workbookXml || !relsXml) throw new Error('ไฟล์นี้ไม่ใช่ Excel มาตรฐาน หรือไฟล์อาจเสีย');
  const refs = parseWorkbookRefs(workbookXml, relsXml);
  const sharedStrings = parseSharedStrings(sharedStringsXml ?? '');
  const sheets: ParsedSheet[] = [];
  const warnings: string[] = [];
  for (const ref of refs) {
    const xml = await zip.file(ref.path)?.async('text');
    if (!xml) {
      warnings.push(`อ่าน sheet ${ref.name} ไม่ได้`);
      continue;
    }
    sheets.push(parseWorksheet(xml, sharedStrings, ref.name));
  }
  if (!sheets.length) throw new Error('ไม่พบ sheet ที่อ่านได้ในไฟล์นี้');
  return { name: file.name, parsedAt: new Date().toISOString(), sheets, warnings };
}

function isRfc(soTypeId: string): boolean {
  return normalize(soTypeId) === 'rfc' || /\bRFC\b/i.test(soTypeId);
}

function keepRfcNegative(value: number, soTypeId: string): number {
  if (!Number.isFinite(value)) return 0;
  return isRfc(soTypeId) && value > 0 ? -value : value;
}

function get(row: RawRow, map: FieldMap, field: FieldKey): Cell {
  const header = map[field];
  return header ? row[header] ?? '' : '';
}

function buildInvoiceLines(sheet: ParsedSheet, map: FieldMap): InvoiceLine[] {
  if (!map.shipQtyPCS || !map.invoiceAmt) return [];
  return sheet.rows.flatMap((row, index) => {
    const soTypeId = safeText(get(row, map, 'soTypeId'));
    const sku = safeText(get(row, map, 'sku'));
    const description = safeText(get(row, map, 'description'));
    const rawQty = parseNumber(get(row, map, 'shipQtyPCS'));
    const rawAmt = parseNumber(get(row, map, 'invoiceAmt'));
    const hasLineIdentity = Boolean(sku || description || safeText(get(row, map, 'sellerGroup')) || safeText(get(row, map, 'brand')));
    const hasLineValues = Number.isFinite(rawQty) || Number.isFinite(rawAmt);
    if (!hasLineIdentity || !hasLineValues) return [];
    return [{
      sourceRowNo: index + 2,
      sellerGroup: safeText(get(row, map, 'sellerGroup')) || 'ไม่ระบุ PS',
      brand: safeText(get(row, map, 'brand')) || 'ไม่ระบุ Brand',
      tasSizeGroup: safeText(get(row, map, 'tasSizeGroup')) || 'ไม่ระบุ Size',
      sku: sku || 'ไม่ระบุ SKU',
      description: description || 'ไม่ระบุ Description',
      soTypeId: soTypeId || 'ไม่ระบุ SOTypeID',
      date: dateText(get(row, map, 'date')) || 'ไม่ระบุวันที่',
      shipQtyPCS: keepRfcNegative(Number.isFinite(rawQty) ? rawQty : 0, soTypeId),
      invoiceAmt: keepRfcNegative(Number.isFinite(rawAmt) ? rawAmt : 0, soTypeId),
      raw,
    } as InvoiceLine];
  });
}

function buildPivot(lines: InvoiceLine[]): PivotRow[] {
  const map = new Map<string, PivotRow>();
  lines.forEach((line) => {
    const keyParts = GROUP_FIELDS.map((field) => line[field]);
    const key = JSON.stringify(keyParts);
    const current = map.get(key) ?? {
      sellerGroup: line.sellerGroup,
      brand: line.brand,
      tasSizeGroup: line.tasSizeGroup,
      sku: line.sku,
      description: line.description,
      soTypeId: line.soTypeId,
      shipQtyPCS: 0,
      invoiceAmt: 0,
      lineCount: 0,
    };
    current.shipQtyPCS += line.shipQtyPCS;
    current.invoiceAmt += line.invoiceAmt;
    current.lineCount += 1;
    map.set(key, current);
  });
  return Array.from(map.values()).sort((a, b) =>
    a.sellerGroup.localeCompare(b.sellerGroup) ||
    a.brand.localeCompare(b.brand) ||
    a.tasSizeGroup.localeCompare(b.tasSizeGroup) ||
    a.sku.localeCompare(b.sku) ||
    a.description.localeCompare(b.description) ||
    a.soTypeId.localeCompare(b.soTypeId)
  );
}

function totalQty(lines: { shipQtyPCS: number }[]): number {
  return lines.reduce((sum, line) => sum + line.shipQtyPCS, 0);
}

function totalAmt(lines: { invoiceAmt: number }[]): number {
  return lines.reduce((sum, line) => sum + line.invoiceAmt, 0);
}

function uniqueOptions(values: string[]): string[] {
  return Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b));
}

function fieldWarnings(headers: string[], map: FieldMap): string[] {
  const normalizedHeaders = headers.map(normalize);
  const warnings: string[] = [];
  REQUIRED_FIELDS.forEach((field) => {
    if (!map[field]) warnings.push(`ไม่พบคอลัมน์จำเป็น: ${FIELD_LABELS[field]}`);
  });
  const ignoredTotals = headers.filter((header) => DISALLOWED_AMOUNT_HEADERS.includes(normalize(header)) || /tot.*invc|invoice.*total|total.*invoice/i.test(header));
  if (ignoredTotals.length) warnings.push(`พบคอลัมน์ยอดหัวบิลที่ถูกเมิน: ${ignoredTotals.join(', ')} — ระบบใช้เฉพาะ InvoiceAmt ต่อบรรทัดเท่านั้น`);
  if (normalizedHeaders.includes('invoiceamt') && !map.invoiceAmt) warnings.push('พบชื่อคล้าย InvoiceAmt แต่จับคู่ไม่ได้ กรุณาตรวจหัวตาราง');
  return warnings;
}

function isPivotReady(map: FieldMap): boolean {
  return REQUIRED_FIELDS.every((field) => Boolean(map[field]));
}

function exportPivotCsv(rows: PivotRow[]) {
  const headers = ['SO Seller Group', 'Brand', 'TAS_SizeGroup', 'SKU', 'Description', 'SOTypeID', 'ShipQtyPCS', 'InvoiceAmt', 'LineCount'];
  const body = rows.map((row) => [row.sellerGroup, row.brand, row.tasSizeGroup, row.sku, row.description, row.soTypeId, row.shipQtyPCS, row.invoiceAmt, row.lineCount].map(csvEscape).join(','));
  downloadText('doit-pivot-result.csv', `\uFEFF${[headers.join(','), ...body].join('\n')}`, 'text/csv;charset=utf-8');
}

function exportReport(workbook: Workbook | null, sheet: ParsedSheet | null, map: FieldMap, filteredLines: InvoiceLine[], pivotRows: PivotRow[], warnings: string[]) {
  if (!workbook || !sheet) return;
  const lines = [
    '# DOIT Pivot Check',
    `ไฟล์: ${workbook.name}`,
    `Sheet: ${sheet.name}`,
    `เวลาที่อ่าน: ${new Date(workbook.parsedAt).toLocaleString('th-TH')}`,
    '',
    '## Field mapping',
    ...REQUIRED_FIELDS.map((field) => `- ${FIELD_LABELS[field]} = ${map[field] ?? 'ไม่พบ'}`),
    '',
    '## กติกาที่ใช้',
    '- source grain = invoice line',
    '- filter raw invoice lines ก่อน aggregate',
    '- group by SO Seller Group, Brand, TAS_SizeGroup, SKU, Description, SOTypeID',
    '- SUM(ShipQtyPCS)',
    '- SUM(InvoiceAmt)',
    '- RFC คงเป็นค่าลบ',
    '- ไม่ใช้ TotInvc / invoice header total',
    '- ไม่กระจายยอดหัวบิลลง SKU',
    '- ไม่สร้าง unit price ปลอม',
    '',
    '## สรุปผล',
    `Raw invoice lines หลัง filter: ${formatNumber(filteredLines.length, 0)}`,
    `Pivot rows: ${formatNumber(pivotRows.length, 0)}`,
    `SUM ShipQtyPCS: ${formatNumber(totalQty(filteredLines))}`,
    `SUM InvoiceAmt: ${formatNumber(totalAmt(filteredLines))}`,
    '',
    '## คำเตือน',
    ...(warnings.length ? warnings.map((w) => `- ${w}`) : ['- ไม่มีคำเตือน']),
  ];
  downloadText('doit-pivot-report.md', lines.join('\n'), 'text/markdown;charset=utf-8');
}

export default function DoitPivotApp() {
  const t = theme();
  const [workbook, setWorkbook] = useState<Workbook | null>(null);
  const [selectedSheetName, setSelectedSheetName] = useState('');
  const [sellerGroupFilter, setSellerGroupFilter] = useState('ALL');
  const [dateFilter, setDateFilter] = useState('ALL');
  const [soTypeFilter, setSoTypeFilter] = useState('ALL');
  const [query, setQuery] = useState('');
  const [status, setStatus] = useState('พร้อมอ่านไฟล์ DOIT');
  const [error, setError] = useState('');

  const activeSheet = useMemo(() => workbook?.sheets.find((sheet) => sheet.name === selectedSheetName) ?? workbook?.sheets[0] ?? null, [workbook, selectedSheetName]);
  const fieldMap = useMemo(() => activeSheet ? detectFields(activeSheet.headers) : {}, [activeSheet]);
  const allLines = useMemo(() => activeSheet ? buildInvoiceLines(activeSheet, fieldMap) : [], [activeSheet, fieldMap]);
  const warnings = useMemo(() => activeSheet ? [...(workbook?.warnings ?? []), ...fieldWarnings(activeSheet.headers, fieldMap)] : workbook?.warnings ?? [], [activeSheet, fieldMap, workbook]);
  const sellerGroups = useMemo(() => uniqueOptions(allLines.map((line) => line.sellerGroup)), [allLines]);
  const dates = useMemo(() => uniqueOptions(allLines.map((line) => line.date)), [allLines]);
  const soTypes = useMemo(() => uniqueOptions(allLines.map((line) => line.soTypeId)), [allLines]);
  const filteredLines = useMemo(() => {
    const q = normalize(query);
    return allLines.filter((line) => {
      if (sellerGroupFilter !== 'ALL' && line.sellerGroup !== sellerGroupFilter) return false;
      if (dateFilter !== 'ALL' && line.date !== dateFilter) return false;
      if (soTypeFilter !== 'ALL' && line.soTypeId !== soTypeFilter) return false;
      if (!q) return true;
      return [line.sellerGroup, line.brand, line.tasSizeGroup, line.sku, line.description, line.soTypeId, line.date].some((value) => normalize(value).includes(q));
    });
  }, [allLines, sellerGroupFilter, dateFilter, soTypeFilter, query]);
  const pivotRows = useMemo(() => buildPivot(filteredLines), [filteredLines]);
  const previewPivotRows = useMemo(() => pivotRows.slice(0, 250), [pivotRows]);
  const previewLines = useMemo(() => filteredLines.slice(0, 80), [filteredLines]);
  const ready = isPivotReady(fieldMap);
  const rfcLines = filteredLines.filter((line) => isRfc(line.soTypeId));

  async function handleFile(file: File) {
    if (file.size > MAX_FILE_SIZE_BYTES) {
      setError(`ไฟล์ใหญ่ ${formatNumber(file.size / 1024 / 1024)} MB เกิน ${MAX_FILE_SIZE_MB} MB`);
      return;
    }
    setError('');
    setStatus(`กำลังอ่านไฟล์ ${file.name}`);
    try {
      const parsed = await parseFile(file);
      const firstPivotSheet = parsed.sheets.find((sheet) => {
        const map = detectFields(sheet.headers);
        return Boolean(map.shipQtyPCS && map.invoiceAmt);
      }) ?? parsed.sheets[0];
      setWorkbook(parsed);
      setSelectedSheetName(firstPivotSheet?.name ?? '');
      setSellerGroupFilter('ALL');
      setDateFilter('ALL');
      setSoTypeFilter('ALL');
      setQuery('');
      setStatus(`อ่านไฟล์สำเร็จ: ${parsed.sheets.length} sheet`);
    } catch (e) {
      setWorkbook(null);
      setSelectedSheetName('');
      setError((e as Error).message || 'อ่านไฟล์ไม่สำเร็จ');
      setStatus('อ่านไฟล์ไม่สำเร็จ');
    }
  }

  const card = { background: t.card, border: `1px solid ${t.border}`, borderRadius: 18, padding: 18 };
  const panel = { background: t.panel, border: `1px solid ${t.border}`, borderRadius: 22, padding: 22, boxShadow: '0 18px 48px rgba(15,23,42,0.06)' };
  const button = { border: 'none', borderRadius: 14, padding: '16px 20px', background: t.accent, color: '#fff', fontWeight: 900, fontSize: 18, cursor: 'pointer' };
  const lightButton = { ...button, background: t.panel, color: t.text, border: `1px solid ${t.border}` };
  const input = { padding: 14, borderRadius: 12, border: `1px solid ${t.border}`, background: '#fff', color: t.text, fontSize: 17 };

  return (
    <div style={{ minHeight: '100vh', background: t.shell, color: t.text, fontFamily: 'Tahoma, Noto Sans Thai, system-ui, sans-serif', fontSize: 18 }}>
      <div style={{ maxWidth: 1500, margin: '0 auto', padding: 22 }}>
        <div style={{ ...panel, display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) auto', gap: 18, alignItems: 'center' }}>
          <div>
            <div style={{ color: t.muted, fontWeight: 900 }}>DOIT Pivot App ใหม่</div>
            <h1 style={{ margin: '8px 0', fontSize: 'clamp(2.2rem, 4vw, 3.8rem)', lineHeight: 1.05 }}>ตรวจยอดตาม Pivot จริง</h1>
            <div style={{ color: t.muted, lineHeight: 1.6 }}>อ่านจาก invoice line โดยตรง แล้ว SUM เฉพาะ ShipQtyPCS และ InvoiceAmt ตาม group ของ DOIT</div>
          </div>
          <label style={{ ...button, display: 'inline-block' }}>
            เลือกไฟล์ DOIT
            <input type="file" accept=".xlsx,.xlsm,.csv,.txt" style={{ display: 'none' }} onChange={(e) => { const file = e.target.files?.[0]; if (file) void handleFile(file); }} />
          </label>
        </div>

        <div style={{ marginTop: 16, ...panel, padding: 18 }}>
          <strong>สถานะ:</strong> {status}
          {workbook && <span style={{ color: t.muted }}> · {workbook.name} · {new Date(workbook.parsedAt).toLocaleString('th-TH')}</span>}
          {error && <div style={{ marginTop: 12, color: t.danger, fontWeight: 900 }}>{error}</div>}
        </div>

        {!workbook && (
          <div style={{ marginTop: 18, display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: 16 }}>
            {[
              ['1', 'เลือกไฟล์ DOIT', 'ใช้ Excel หรือ CSV จาก DOIT'],
              ['2', 'ระบบอ่าน raw invoice line', 'ยังไม่รวมยอด ยังไม่แตะ TotInvc'],
              ['3', 'เลือก PS/วันที่', 'filter raw lines ก่อนเสมอ'],
              ['4', 'ดู Pivot', 'group แล้ว SUM เหมือน Pivot'],
            ].map(([no, title, desc]) => <div key={no} style={card}><div style={{ fontSize: 34, fontWeight: 900, color: t.accent }}>{no}</div><strong>{title}</strong><div style={{ color: t.muted, marginTop: 6 }}>{desc}</div></div>)}
          </div>
        )}

        {workbook && activeSheet && (
          <>
            <div style={{ marginTop: 18, display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(210px, 1fr))', gap: 14 }}>
              {[
                ['Raw lines หลัง filter', filteredLines.length],
                ['Pivot rows', pivotRows.length],
                ['SUM ShipQtyPCS', totalQty(filteredLines)],
                ['SUM InvoiceAmt', totalAmt(filteredLines)],
                ['RFC amount', totalAmt(rfcLines)],
              ].map(([label, value]) => <div key={String(label)} style={card}><div style={{ color: t.muted, fontSize: 15 }}>{label}</div><div style={{ fontSize: 32, fontWeight: 900, marginTop: 4 }}>{typeof value === 'number' ? formatNumber(value) : value}</div></div>)}
            </div>

            <div style={{ marginTop: 18, ...panel }}>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(190px, 1fr))', gap: 12 }}>
                <label style={{ display: 'grid', gap: 6 }}>
                  <span style={{ color: t.muted }}>Sheet</span>
                  <select value={activeSheet.name} onChange={(e) => setSelectedSheetName(e.target.value)} style={input}>
                    {workbook.sheets.map((sheet) => <option key={sheet.name} value={sheet.name}>{sheet.name}</option>)}
                  </select>
                </label>
                <label style={{ display: 'grid', gap: 6 }}>
                  <span style={{ color: t.muted }}>PS / SO Seller Group</span>
                  <select value={sellerGroupFilter} onChange={(e) => setSellerGroupFilter(e.target.value)} style={input}>
                    <option value="ALL">ทุก PS</option>
                    {sellerGroups.map((value) => <option key={value} value={value}>{value}</option>)}
                  </select>
                </label>
                <label style={{ display: 'grid', gap: 6 }}>
                  <span style={{ color: t.muted }}>วันที่</span>
                  <select value={dateFilter} onChange={(e) => setDateFilter(e.target.value)} style={input}>
                    <option value="ALL">ทุกวันที่พบ</option>
                    {dates.map((value) => <option key={value} value={value}>{value}</option>)}
                  </select>
                </label>
                <label style={{ display: 'grid', gap: 6 }}>
                  <span style={{ color: t.muted }}>SOTypeID</span>
                  <select value={soTypeFilter} onChange={(e) => setSoTypeFilter(e.target.value)} style={input}>
                    <option value="ALL">ทุก SOTypeID</option>
                    {soTypes.map((value) => <option key={value} value={value}>{value}</option>)}
                  </select>
                </label>
                <label style={{ display: 'grid', gap: 6 }}>
                  <span style={{ color: t.muted }}>ค้นหา</span>
                  <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="SKU / Brand / Description" style={input} />
                </label>
              </div>
              <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 16 }}>
                <button style={button} onClick={() => exportPivotCsv(pivotRows)} disabled={!pivotRows.length}>โหลด Pivot CSV</button>
                <button style={lightButton} onClick={() => exportReport(workbook, activeSheet, fieldMap, filteredLines, pivotRows, warnings)}>โหลดรายงาน</button>
                <button style={lightButton} onClick={() => { setWorkbook(null); setSelectedSheetName(''); setStatus('พร้อมอ่านไฟล์ DOIT'); }}>เริ่มใหม่</button>
              </div>
            </div>

            <div style={{ marginTop: 18, ...panel }}>
              <h2 style={{ marginTop: 0 }}>ตรวจ field ที่ใช้คำนวณ</h2>
              <div style={{ color: ready ? t.good : t.danger, fontWeight: 900, marginBottom: 12 }}>{ready ? 'พร้อมคำนวณแบบ Pivot' : 'ยังขาด field สำคัญ ไม่ควรเชื่อยอดจนกว่าจะใช้ไฟล์ DOIT ที่มี field ครบ'}</div>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: 10 }}>
                {REQUIRED_FIELDS.map((field) => <div key={field} style={card}><div style={{ color: t.muted }}>{FIELD_LABELS[field]}</div><strong>{fieldMap[field] ?? 'ไม่พบ'}</strong></div>)}
              </div>
              {warnings.length > 0 && <div style={{ marginTop: 14, display: 'grid', gap: 8 }}>{warnings.map((warning, index) => <div key={index} style={{ padding: 12, borderRadius: 12, background: '#fef2f2', color: t.danger, fontWeight: 800 }}>{warning}</div>)}</div>}
            </div>

            <div style={{ marginTop: 18, ...panel }}>
              <h2 style={{ marginTop: 0 }}>Pivot result</h2>
              <div style={{ color: t.muted, marginBottom: 12 }}>group by: SO Seller Group, Brand, TAS_SizeGroup, SKU, Description, SOTypeID · แสดง 250 แถวแรก</div>
              <div style={{ overflowX: 'auto', border: `1px solid ${t.border}`, borderRadius: 16 }}>
                <table style={{ width: '100%', minWidth: 1120, borderCollapse: 'collapse' }}>
                  <thead><tr>{['SO Seller Group', 'Brand', 'TAS_SizeGroup', 'SKU', 'Description', 'SOTypeID', 'ShipQtyPCS', 'InvoiceAmt', 'LineCount'].map((head) => <th key={head} style={{ textAlign: 'left', padding: 12, borderBottom: `1px solid ${t.border}`, background: t.card }}>{head}</th>)}</tr></thead>
                  <tbody>{previewPivotRows.map((row, index) => <tr key={`${row.sellerGroup}-${row.brand}-${row.sku}-${row.soTypeId}-${index}`} style={{ background: index % 2 ? t.tableAlt : '#fff' }}>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{row.sellerGroup}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{row.brand}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{row.tasSizeGroup}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{row.sku}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{row.description}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{row.soTypeId}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}`, textAlign: 'right' }}>{formatNumber(row.shipQtyPCS)}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}`, textAlign: 'right' }}>{formatNumber(row.invoiceAmt)}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}`, textAlign: 'right' }}>{formatNumber(row.lineCount, 0)}</td>
                  </tr>)}</tbody>
                </table>
              </div>
            </div>

            <div style={{ marginTop: 18, ...panel }}>
              <h2 style={{ marginTop: 0 }}>Raw invoice lines หลัง filter</h2>
              <div style={{ color: t.muted, marginBottom: 12 }}>แสดง 80 แถวแรก ใช้ตรวจว่ากรองก่อน group ถูกต้อง</div>
              <div style={{ overflowX: 'auto', border: `1px solid ${t.border}`, borderRadius: 16 }}>
                <table style={{ width: '100%', minWidth: 1120, borderCollapse: 'collapse' }}>
                  <thead><tr>{['Row', 'Date', 'SO Seller Group', 'Brand', 'TAS_SizeGroup', 'SKU', 'Description', 'SOTypeID', 'ShipQtyPCS', 'InvoiceAmt'].map((head) => <th key={head} style={{ textAlign: 'left', padding: 12, borderBottom: `1px solid ${t.border}`, background: t.card }}>{head}</th>)}</tr></thead>
                  <tbody>{previewLines.map((line, index) => <tr key={`${line.sourceRowNo}-${index}`} style={{ background: index % 2 ? t.tableAlt : '#fff' }}>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{line.sourceRowNo}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{line.date}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{line.sellerGroup}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{line.brand}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{line.tasSizeGroup}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{line.sku}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{line.description}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{line.soTypeId}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}`, textAlign: 'right' }}>{formatNumber(line.shipQtyPCS)}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${t.border}`, textAlign: 'right' }}>{formatNumber(line.invoiceAmt)}</td>
                  </tr>)}</tbody>
                </table>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
