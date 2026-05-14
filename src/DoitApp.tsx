import JSZip from 'jszip';
import { useEffect, useMemo, useState } from 'react';

type CellValue = string | number | boolean | null;
type RowObject = Record<string, CellValue>;

type NumericStat = { name: string; total: number; count: number };
type SuspiciousInvoice = { invoiceNo: string; rows: number; amount: number; skus: number; sheet: string };

type SheetData = {
  name: string;
  headers: string[];
  rows: RowObject[];
  rawRowCount: number;
  numericSummary: NumericStat[];
  suspicious: SuspiciousInvoice[];
};

type WorkbookData = {
  sourceName: string;
  parsedAt: string;
  sheets: SheetData[];
  warnings: string[];
  warningCount: number;
  inflationHits: number;
  rowCount: number;
};

type HistoryEntry = {
  sourceName: string;
  parsedAt: string;
  sheets: number;
  rows: number;
  warnings: number;
  inflationHits: number;
};

type View = 'summary' | 'sheets' | 'history' | 'settings';

type Theme = {
  shell: string;
  panel: string;
  card: string;
  border: string;
  text: string;
  muted: string;
  accent: string;
  danger: string;
  success: string;
  tableAlt: string;
};

const SETTINGS_KEY = 'doit.web.settings';
const HISTORY_KEY = 'doit.web.history';

function readJson<T>(key: string, fallback: T): T {
  if (typeof window === 'undefined') return fallback;
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback;
  }
}

function writeJson(key: string, value: unknown) {
  if (typeof window === 'undefined') return;
  window.localStorage.setItem(key, JSON.stringify(value));
}

function cleanText(value: unknown): string {
  return String(value ?? '')
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function compact(value: unknown): string {
  return cleanText(value)
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[\s\-_\/().,|]+/g, '')
    .replace(/[^\p{L}\p{N}]+/gu, '');
}

function colToIndex(ref: string): number {
  const match = /([A-Z]+)/i.exec(ref);
  if (!match) return 0;
  return match[1].toUpperCase().split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
}

function indexToCol(index: number): string {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - mod) / 26);
  }
  return s;
}

function xmlText(node: Element | null | undefined): string {
  return cleanText(node?.textContent ?? '');
}

function splitCsvLine(line: string): string[] {
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
  return out.map((cell) => cell.trim());
}

function findHeader(headers: string[], patterns: RegExp[], reject?: (header: string) => boolean): string | undefined {
  return headers.find((header) => {
    if (reject?.(header)) return false;
    const normalized = cleanText(header).toLowerCase();
    return patterns.some((pattern) => pattern.test(normalized));
  });
}

const HEADER_TOTAL_PATTERNS = [
  'totinvc',
  'totinvcamt',
  'totalinvoice',
  'totalinvoiceamount',
  'grandtotal',
  'grand_total',
  'totalamount',
  'total_amount',
  'sumofinvoiceamt',
  'sumofamount',
  'ยอดรวม',
  'ยอดรวมทั้งสิ้น',
];

function isHeaderTotalColumn(name: string): boolean {
  const f = compact(name);
  if (!f) return false;
  if (HEADER_TOTAL_PATTERNS.some((pattern) => f === compact(pattern))) return true;
  if (f.includes('grandtotal') || f.includes('grandtotalamount')) return true;
  if (f.startsWith('totinvc') || f.startsWith('totalinvc')) return true;
  if (f === 'total' || f === 'totalamount' || f === 'totalsum') return true;
  return false;
}

function parseSharedStrings(xml: string): string[] {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return Array.from(doc.getElementsByTagName('si')).map((item) => cleanText(item.textContent ?? ''));
}

function parseWorkbookSheets(workbookXml: string, relsXml: string): { name: string; path: string }[] {
  const workbookDoc = new DOMParser().parseFromString(workbookXml, 'application/xml');
  const relsDoc = new DOMParser().parseFromString(relsXml, 'application/xml');
  const relMap = new Map<string, string>();

  Array.from(relsDoc.getElementsByTagName('Relationship')).forEach((rel) => {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (id && target) relMap.set(id, target.startsWith('/') ? target.slice(1) : target);
  });

  return Array.from(workbookDoc.getElementsByTagName('sheet')).map((sheet) => {
    const name = sheet.getAttribute('name') ?? 'Sheet';
    const rid = sheet.getAttribute('r:id') ?? sheet.getAttribute('id') ?? '';
    const target = relMap.get(rid) ?? 'worksheets/sheet1.xml';
    const normalized = target.startsWith('xl/') ? target : `xl/${target}`;
    return { name, path: normalized };
  });
}

function parseWorksheet(xml: string, sharedStrings: string[], sheetName: string): SheetData {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const rows = Array.from(doc.getElementsByTagName('row'));
  const matrix: Array<Array<CellValue>> = [];

  for (const row of rows) {
    const rowIndex = Number(row.getAttribute('r') ?? matrix.length + 1) - 1;
    if (!matrix[rowIndex]) matrix[rowIndex] = [];

    for (const cell of Array.from(row.getElementsByTagName('c'))) {
      const ref = cell.getAttribute('r') ?? '';
      const colIndex = ref ? colToIndex(ref) : matrix[rowIndex].length;
      const type = cell.getAttribute('t') ?? '';
      const valueNode = cell.getElementsByTagName('v')[0];
      const inlineNode = cell.getElementsByTagName('is')[0];
      let value: CellValue = null;

      if (type === 's') {
        const idx = Number(xmlText(valueNode));
        value = sharedStrings[idx] ?? '';
      } else if (type === 'inlineStr') {
        value = cleanText(inlineNode?.textContent ?? '');
      } else if (type === 'b') {
        value = xmlText(valueNode) === '1';
      } else {
        const raw = xmlText(valueNode);
        if (raw === '') value = '';
        else if (!Number.isNaN(Number(raw))) value = Number(raw);
        else value = raw;
      }

      matrix[rowIndex][colIndex] = value;
    }
  }

  const firstNonEmptyRow = matrix.findIndex((row) => row?.some((value) => cleanText(value).length > 0));
  const headerRow = firstNonEmptyRow >= 0 ? matrix[firstNonEmptyRow] : [];
  const headers = (headerRow ?? []).map((value, i) => cleanText(value) || `Column ${indexToCol(i)}`);

  const bodyRows = matrix
    .slice(firstNonEmptyRow + 1)
    .filter((row) => row?.some((value) => cleanText(value).length > 0));

  const rowsData: RowObject[] = bodyRows.map((row) => {
    const obj: RowObject = {};
    headers.forEach((header, i) => {
      obj[header] = row?.[i] ?? '';
    });
    return obj;
  });

  return {
    name: sheetName,
    headers,
    rows: rowsData,
    rawRowCount: rowsData.length,
    numericSummary: [],
    suspicious: [],
  };
}

function parseCsv(text: string, sheetName: string): SheetData {
  const lines = text.split(/\r?\n/).filter((line) => cleanText(line).length > 0);
  const headers = splitCsvLine(lines[0] ?? '').map((h, i) => cleanText(h) || `Column ${indexToCol(i)}`);
  const rows: RowObject[] = lines.slice(1).map((line) => {
    const cells = splitCsvLine(line);
    const row: RowObject = {};
    headers.forEach((header, i) => {
      row[header] = cells[i] ?? '';
    });
    return row;
  });

  return {
    name: sheetName,
    headers,
    rows,
    rawRowCount: rows.length,
    numericSummary: [],
    suspicious: [],
  };
}

function sumNumericColumns(rows: RowObject[]): NumericStat[] {
  const totals = new Map<string, { total: number; count: number }>();
  for (const row of rows) {
    for (const [key, value] of Object.entries(row)) {
      const n = typeof value === 'number' ? value : Number(String(value).replace(/,/g, ''));
      if (!Number.isFinite(n)) continue;
      const current = totals.get(key) ?? { total: 0, count: 0 };
      current.total += n;
      current.count += 1;
      totals.set(key, current);
    }
  }

  return Array.from(totals.entries())
    .map(([name, value]) => ({ name, total: value.total, count: value.count }))
    .sort((a, b) => Math.abs(b.total) - Math.abs(a.total))
    .slice(0, 8);
}

function detectSuspiciousInvoices(rows: RowObject[], headers: string[], sheetName: string): SuspiciousInvoice[] {
  const invoiceKey = findHeader(headers, [/invoice.*no/i, /invoice_no/i, /invoiceno/i, /inv\.?\s*no/i, /bill.*no/i])
    ?? findHeader(headers, [/invoice/i]);
  const amountKey = findHeader(headers, [/invoice.*amt/i, /\bamt\b/i, /\bamount\b/i, /\bprice\b/i, /\bvalue\b/i, /\bยอด\b/i, /\bเงิน\b/i], isHeaderTotalColumn);
  const skuKey = findHeader(headers, [/sku/i, /item/i, /product/i, /code/i, /desc/i]);
  if (!invoiceKey || !amountKey) return [];

  const grouped = new Map<string, { amounts: Set<number>; skus: Set<string>; rows: number }>();
  for (const row of rows) {
    const invoiceNo = cleanText(row[invoiceKey]);
    if (!invoiceNo) continue;
    const amountValue = row[amountKey];
    const amount = typeof amountValue === 'number' ? amountValue : Number(String(amountValue).replace(/,/g, ''));
    if (!Number.isFinite(amount) || amount <= 0) continue;
    const sku = skuKey ? cleanText(row[skuKey]) : '';

    const current = grouped.get(invoiceNo) ?? { amounts: new Set<number>(), skus: new Set<string>(), rows: 0 };
    current.amounts.add(amount);
    if (sku) current.skus.add(sku);
    current.rows += 1;
    grouped.set(invoiceNo, current);
  }

  const suspicious: SuspiciousInvoice[] = [];
  for (const [invoiceNo, group] of grouped.entries()) {
    if (group.rows >= 2 && group.amounts.size === 1 && group.skus.size >= 2) {
      suspicious.push({
        invoiceNo,
        rows: group.rows,
        amount: Array.from(group.amounts)[0] ?? 0,
        skus: group.skus.size,
        sheet: sheetName,
      });
    }
  }

  return suspicious.sort((a, b) => b.rows - a.rows || b.amount - a.amount).slice(0, 12);
}

function downloadText(filename: string, content: string) {
  const blob = new Blob([content], { type: 'application/json;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

async function parseWorkbookFile(file: File): Promise<WorkbookData> {
  const warnings: string[] = [];
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const workbookXml = await zip.file('xl/workbook.xml')?.async('text');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('text');
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('text');

  if (!workbookXml || !relsXml) {
    throw new Error('ไฟล์นี้ไม่ใช่ workbook มาตรฐาน');
  }

  const sharedStrings = parseSharedStrings(sharedStringsXml ?? '');
  const sheetDefs = parseWorkbookSheets(workbookXml, relsXml);
  const sheets: SheetData[] = [];

  for (const sheetDef of sheetDefs) {
    const xml = await zip.file(sheetDef.path)?.async('text');
    if (!xml) {
      warnings.push(`ไม่พบข้อมูลของชีต ${sheetDef.name}`);
      continue;
    }

    const parsed = parseWorksheet(xml, sharedStrings, sheetDef.name);
    parsed.numericSummary = sumNumericColumns(parsed.rows);
    parsed.suspicious = detectSuspiciousInvoices(parsed.rows, parsed.headers, parsed.name);
    sheets.push(parsed);
  }

  if (!sheets.length) {
    throw new Error('ไม่พบข้อมูลที่ใช้งานได้ในไฟล์');
  }

  const inflationHits = sheets.reduce((sum, sheet) => sum + sheet.suspicious.length, 0);
  const rowCount = sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);

  return {
    sourceName: file.name,
    parsedAt: new Date().toISOString(),
    sheets,
    warnings,
    warningCount: warnings.length,
    inflationHits,
    rowCount,
  };
}

async function parseTextFile(file: File): Promise<WorkbookData> {
  const text = await file.text();
  const sheet = parseCsv(text, file.name);
  sheet.numericSummary = sumNumericColumns(sheet.rows);
  sheet.suspicious = detectSuspiciousInvoices(sheet.rows, sheet.headers, sheet.name);
  return {
    sourceName: file.name,
    parsedAt: new Date().toISOString(),
    sheets: [sheet],
    warnings: ['ใช้ตัวอ่านข้อความแบบ fallback'],
    warningCount: 1,
    inflationHits: sheet.suspicious.length,
    rowCount: sheet.rows.length,
  };
}

function themeStyles(darkMode: boolean): Theme {
  return darkMode
    ? {
        shell: '#0f172a',
        panel: '#111827',
        card: '#0b1220',
        border: '#334155',
        text: '#e2e8f0',
        muted: '#94a3b8',
        accent: '#38bdf8',
        danger: '#fb7185',
        success: '#34d399',
        tableAlt: 'rgba(255,255,255,0.02)',
      }
    : {
        shell: '#f8fafc',
        panel: '#ffffff',
        card: '#f1f5f9',
        border: '#cbd5e1',
        text: '#0f172a',
        muted: '#475569',
        accent: '#0ea5e9',
        danger: '#e11d48',
        success: '#059669',
        tableAlt: 'rgba(15,23,42,0.03)',
      };
}

export default function DoitApp() {
  const settings = readJson<{ darkMode: boolean; rowsLimit: number }>(SETTINGS_KEY, { darkMode: true, rowsLimit: 75 });
  const [darkMode, setDarkMode] = useState(settings.darkMode);
  const [rowsLimit, setRowsLimit] = useState(settings.rowsLimit);
  const [view, setView] = useState<View>('summary');
  const [status, setStatus] = useState('พร้อมใช้งาน');
  const [error, setError] = useState('');
  const [dragOver, setDragOver] = useState(false);
  const [workbook, setWorkbook] = useState<WorkbookData | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);
  const [query, setQuery] = useState('');
  const [onlySuspicious, setOnlySuspicious] = useState(false);
  const [history, setHistory] = useState<HistoryEntry[]>(() => readJson<HistoryEntry[]>(HISTORY_KEY, []));

  useEffect(() => {
    writeJson(SETTINGS_KEY, { darkMode, rowsLimit });
    if (typeof document !== 'undefined') {
      document.documentElement.style.colorScheme = darkMode ? 'dark' : 'light';
      document.body.style.background = themeStyles(darkMode).shell;
      document.body.style.margin = '0';
    }
  }, [darkMode, rowsLimit]);

  useEffect(() => {
    writeJson(HISTORY_KEY, history.slice(0, 12));
  }, [history]);

  const theme = themeStyles(darkMode);
  const active = workbook?.sheets[activeSheet] ?? null;
  const activeInvoiceKey = active ? findHeader(active.headers, [/invoice.*no/i, /invoice_no/i, /invoiceno/i, /inv\.?\s*no/i, /bill.*no/i]) ?? findHeader(active.headers, [/invoice/i]) : undefined;
  const suspiciousSet = useMemo(() => new Set((active?.suspicious ?? []).map((item) => item.invoiceNo)), [active]);
  const activeRows = useMemo(() => {
    if (!active) return [] as RowObject[];
    let rows = active.rows;
    if (query.trim()) {
      const q = compact(query);
      rows = rows.filter((row) => Object.entries(row).some(([key, value]) => compact(key).includes(q) || compact(value).includes(q)));
    }
    if (onlySuspicious && activeInvoiceKey) {
      rows = rows.filter((row) => suspiciousSet.has(cleanText(row[activeInvoiceKey!]))) ;
    }
    return rows;
  }, [active, query, onlySuspicious, activeInvoiceKey, suspiciousSet]);
  const tableRows = activeRows.slice(0, rowsLimit);
  const numericSummary = active?.numericSummary ?? [];
  const workbookSuspicious = workbook ? workbook.sheets.flatMap((sheet) => sheet.suspicious) : [];
  const visibleWarnings = workbook?.warnings ?? [];

  async function loadFile(file: File) {
    setError('');
    setStatus(`กำลังอ่าน ${file.name}`);
    try {
      const data = /\.(xlsx|xlsm|zip)$/i.test(file.name) ? await parseWorkbookFile(file) : await parseTextFile(file);
      setWorkbook(data);
      setActiveSheet(0);
      setView('summary');
      setStatus(`โหลดสำเร็จ: ${data.sheets.length} sheet`);
      setHistory((prev) => [{
        sourceName: data.sourceName,
        parsedAt: data.parsedAt,
        sheets: data.sheets.length,
        rows: data.rowCount,
        warnings: data.warningCount,
        inflationHits: data.inflationHits,
      }, ...prev].slice(0, 12));
    } catch (e) {
      setWorkbook(null);
      setError((e as Error).message || 'ไม่สามารถอ่านไฟล์ได้');
      setStatus('เกิดข้อผิดพลาด');
    }
  }

  function clearCurrent() {
    setWorkbook(null);
    setActiveSheet(0);
    setQuery('');
    setOnlySuspicious(false);
    setStatus('พร้อมใช้งาน');
    setError('');
  }

  function exportCurrentJson() {
    if (!workbook) return;
    downloadText(
      `doit-${workbook.sourceName.replace(/\.[^.]+$/, '')}.json`,
      JSON.stringify(workbook, null, 2),
    );
  }

  function exportActiveSheetCsv() {
    if (!active) return;
    const headers = active.headers;
    const lines = [headers.join(',')];
    for (const row of activeRows) {
      lines.push(headers.map((header) => {
        const value = cleanText(row[header]);
        return /[",\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
      }).join(','));
    }
    downloadText(`${active.name.replace(/\W+/g, '_')}.csv`, lines.join('\n'));
  }

  const shellStyle: React.CSSProperties = {
    minHeight: '100vh',
    background: theme.shell,
    color: theme.text,
    fontFamily: 'Inter, system-ui, sans-serif',
  };

  const panelStyle: React.CSSProperties = {
    background: theme.panel,
    border: `1px solid ${theme.border}`,
    borderRadius: 20,
    boxShadow: darkMode ? '0 20px 60px rgba(0,0,0,0.18)' : '0 18px 48px rgba(15,23,42,0.06)',
  };

  const cardStyle: React.CSSProperties = {
    background: theme.card,
    border: `1px solid ${theme.border}`,
    borderRadius: 18,
  };

  const navBtn = (id: View, label: string) => ({
    cursor: 'pointer',
    borderRadius: 999,
    padding: '10px 14px',
    border: `1px solid ${view === id ? theme.accent : theme.border}`,
    background: view === id ? `${theme.accent}22` : theme.panel,
    color: theme.text,
    fontWeight: 700 as const,
  });

  const totalRows = workbook?.rowCount ?? 0;
  const activeSuspicious = active?.suspicious ?? [];
  const rowsShown = activeRows.length;

  return (
    <div style={shellStyle}>
      <div style={{ maxWidth: 1400, margin: '0 auto', padding: 24 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 16, alignItems: 'flex-end', flexWrap: 'wrap' }}>
          <div>
            <div style={{ fontSize: 13, letterSpacing: 1.2, textTransform: 'uppercase', color: theme.muted }}>DOIT Web App</div>
            <h1 style={{ margin: '8px 0 6px', fontSize: 'clamp(2rem, 3vw, 3.2rem)', lineHeight: 1.05 }}>Workbook Analyzer</h1>
            <div style={{ color: theme.muted, maxWidth: 820 }}>
              อัปโหลดไฟล์ Excel แล้วอ่านข้อมูลจริงจากหลายชีต, ดูสรุปตัวเลข, ตรวจ invoice ที่มีความเสี่ยงซ้ำซ้อน และ export ผลลัพธ์ได้ทันที
            </div>
          </div>
          <div style={{ ...panelStyle, minWidth: 280, padding: 16 }}>
            <div style={{ fontSize: 12, color: theme.muted }}>สถานะ</div>
            <div style={{ fontWeight: 800, marginTop: 4 }}>{status}</div>
            {workbook && (
              <div style={{ marginTop: 6, fontSize: 12, color: theme.muted }}>
                {workbook.sourceName} · {workbook.sheets.length} sheet · {new Date(workbook.parsedAt).toLocaleString()}
              </div>
            )}
          </div>
        </div>

        <div
          onDragOver={(e) => {
            e.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragOver(false);
            const file = e.dataTransfer.files?.[0];
            if (file) void loadFile(file);
          }}
          style={{
            marginTop: 22,
            padding: 22,
            borderRadius: 22,
            border: `1px dashed ${dragOver ? theme.accent : theme.border}`,
            background: dragOver ? `${theme.accent}12` : theme.panel,
            ...panelStyle,
          }}
        >
          <div style={{ display: 'flex', justifyContent: 'space-between', gap: 16, alignItems: 'center', flexWrap: 'wrap' }}>
            <div>
              <div style={{ fontWeight: 800, fontSize: 18 }}>นำเข้าไฟล์</div>
              <div style={{ color: theme.muted, marginTop: 4 }}>รองรับ .xlsx, .xlsm, .zip, .csv และ .txt</div>
            </div>
            <label style={{ cursor: 'pointer', padding: '12px 16px', borderRadius: 14, background: theme.accent, color: '#082f49', fontWeight: 800 }}>
              เลือกไฟล์
              <input
                type="file"
                accept=".xlsx,.xlsm,.zip,.csv,.txt"
                style={{ display: 'none' }}
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) void loadFile(file);
                }}
              />
            </label>
          </div>
          {error && <div style={{ marginTop: 16, padding: 12, borderRadius: 12, background: `${theme.danger}18`, color: theme.danger }}>{error}</div>}
        </div>

        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 18 }}>
          <button style={navBtn('summary', 'Summary')} onClick={() => setView('summary')}>Summary</button>
          <button style={navBtn('sheets', 'Sheets')} onClick={() => setView('sheets')}>Sheets</button>
          <button style={navBtn('history', 'History')} onClick={() => setView('history')}>History</button>
          <button style={navBtn('settings', 'Settings')} onClick={() => setView('settings')}>Settings</button>
          <button style={navBtn('summary', 'Export JSON')} onClick={exportCurrentJson} disabled={!workbook}>Export JSON</button>
          <button style={navBtn('summary', 'Export CSV')} onClick={exportActiveSheetCsv} disabled={!active}>Export CSV</button>
          <button style={navBtn('summary', 'Clear')} onClick={clearCurrent}>Clear</button>
        </div>

        {workbook && view === 'summary' && (
          <div style={{ marginTop: 24, display: 'grid', gap: 18 }}>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 16 }}>
              {[
                { label: 'Sheets', value: workbook.sheets.length },
                { label: 'Rows', value: totalRows },
                { label: 'Warnings', value: workbook.warningCount },
                { label: 'Suspicious invoices', value: workbook.inflationHits },
              ].map((stat) => (
                <div key={stat.label} style={{ ...cardStyle, padding: 18 }}>
                  <div style={{ fontSize: 12, color: theme.muted }}>{stat.label}</div>
                  <div style={{ fontSize: 28, fontWeight: 900, marginTop: 6 }}>{stat.value}</div>
                </div>
              ))}
            </div>

            <div style={{ ...panelStyle, padding: 18 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
                <div>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>Warnings</div>
                  <div style={{ color: theme.muted, marginTop: 4 }}>สิ่งที่ตัวอ่านพบจากไฟล์ต้นฉบับ</div>
                </div>
                <div style={{ color: theme.muted, fontSize: 13 }}>{visibleWarnings.length} item(s)</div>
              </div>
              <div style={{ marginTop: 14, display: 'grid', gap: 10 }}>
                {visibleWarnings.length ? visibleWarnings.map((warning, i) => (
                  <div key={i} style={{ padding: 12, borderRadius: 12, border: `1px solid ${theme.border}`, background: theme.card }}>{warning}</div>
                )) : <div style={{ color: theme.muted }}>ไม่มี warning</div>}
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: 16 }}>
              {workbook.sheets.map((sheet, idx) => (
                <button
                  key={sheet.name + idx}
                  onClick={() => { setActiveSheet(idx); setView('sheets'); }}
                  style={{
                    textAlign: 'left',
                    padding: 18,
                    borderRadius: 18,
                    border: `1px solid ${theme.border}`,
                    background: theme.panel,
                    color: theme.text,
                    cursor: 'pointer',
                  }}
                >
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}>
                    <div style={{ fontWeight: 800 }}>{sheet.name}</div>
                    <div style={{ color: theme.muted }}>{sheet.rows.length} rows</div>
                  </div>
                  <div style={{ color: theme.muted, marginTop: 6 }}>{sheet.headers.length} columns · {sheet.suspicious.length} suspicious invoice(s)</div>
                </button>
              ))}
            </div>

            <div style={{ ...panelStyle, padding: 18 }}>
              <div style={{ fontWeight: 800, fontSize: 18 }}>Suspicious invoices</div>
              <div style={{ color: theme.muted, marginTop: 4 }}>รายการ invoice ที่ amount ซ้ำหลายบรรทัดพร้อม SKU ต่างกัน</div>
              <div style={{ marginTop: 14, display: 'grid', gap: 10 }}>
                {workbookSuspicious.length ? workbookSuspicious.slice(0, 10).map((item) => (
                  <div key={`${item.sheet}-${item.invoiceNo}`} style={{ display: 'grid', gridTemplateColumns: '1.2fr 0.6fr 0.6fr 0.6fr 0.8fr', gap: 10, padding: 12, borderRadius: 12, background: theme.card, border: `1px solid ${theme.border}` }}>
                    <div><strong>{item.invoiceNo}</strong><div style={{ color: theme.muted, fontSize: 12 }}>{item.sheet}</div></div>
                    <div>{item.rows} rows</div>
                    <div>{item.skus} SKUs</div>
                    <div>{Number.isInteger(item.amount) ? item.amount.toLocaleString() : item.amount.toFixed(2)}</div>
                    <div style={{ color: theme.danger, fontWeight: 800 }}>flagged</div>
                  </div>
                )) : <div style={{ color: theme.muted }}>ยังไม่พบ invoice ที่เข้าข่าย</div>}
              </div>
            </div>
          </div>
        )}

        {workbook && view === 'sheets' && active && (
          <div style={{ marginTop: 24, display: 'grid', gap: 18 }}>
            <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
              {workbook.sheets.map((sheet, idx) => (
                <button
                  key={sheet.name + idx}
                  onClick={() => setActiveSheet(idx)}
                  style={{
                    padding: '10px 14px',
                    borderRadius: 999,
                    border: `1px solid ${idx === activeSheet ? theme.accent : theme.border}`,
                    background: idx === activeSheet ? `${theme.accent}22` : theme.panel,
                    color: theme.text,
                    cursor: 'pointer',
                    fontWeight: 700,
                  }}
                >
                  {sheet.name}
                </button>
              ))}
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1.65fr) minmax(280px, 0.9fr)', gap: 18 }}>
              <div style={{ minWidth: 0 }}>
                <div style={{ ...panelStyle, padding: 18 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                    <div>
                      <div style={{ fontSize: 12, color: theme.muted }}>Active sheet</div>
                      <div style={{ fontSize: 20, fontWeight: 800 }}>{active.name}</div>
                    </div>
                    <div style={{ color: theme.muted }}>{active.rows.length} rows · {active.headers.length} columns</div>
                  </div>
                  <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 14 }}>
                    <input
                      value={query}
                      onChange={(e) => setQuery(e.target.value)}
                      placeholder="Search any cell..."
                      style={{
                        flex: '1 1 260px',
                        padding: '12px 14px',
                        borderRadius: 12,
                        border: `1px solid ${theme.border}`,
                        background: theme.card,
                        color: theme.text,
                      }}
                    />
                    <button
                      onClick={() => setOnlySuspicious((value) => !value)}
                      style={{
                        padding: '12px 14px',
                        borderRadius: 12,
                        border: `1px solid ${onlySuspicious ? theme.accent : theme.border}`,
                        background: onlySuspicious ? `${theme.accent}22` : theme.card,
                        color: theme.text,
                        cursor: 'pointer',
                        fontWeight: 700,
                      }}
                    >
                      {onlySuspicious ? 'Showing suspicious only' : 'Show suspicious only'}
                    </button>
                  </div>
                  <div style={{ marginTop: 12, color: theme.muted }}>
                    Showing {Math.min(rowsLimit, rowsShown)} of {rowsShown} filtered rows
                  </div>
                </div>

                <div style={{ marginTop: 18, overflowX: 'auto', borderRadius: 18, border: `1px solid ${theme.border}`, background: theme.panel }}>
                  <table style={{ width: '100%', minWidth: 900, borderCollapse: 'collapse' }}>
                    <thead>
                      <tr>
                        {active.headers.map((header) => (
                          <th
                            key={header}
                            style={{
                              position: 'sticky',
                              top: 0,
                              background: theme.panel,
                              color: theme.muted,
                              padding: 12,
                              textAlign: 'left',
                              borderBottom: `1px solid ${theme.border}`,
                              whiteSpace: 'nowrap',
                            }}
                          >
                            {header}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {tableRows.map((row, idx) => {
                        const invoiceNo = activeInvoiceKey ? cleanText(row[activeInvoiceKey]) : '';
                        const suspicious = invoiceNo && suspiciousSet.has(invoiceNo);
                        return (
                          <tr key={idx} style={{ background: suspicious ? `${theme.danger}10` : idx % 2 ? theme.tableAlt : 'transparent' }}>
                            {active.headers.map((header) => (
                              <td key={header} style={{ padding: 12, borderBottom: `1px solid ${theme.border}`, verticalAlign: 'top' }}>
                                {cleanText(row[header]) || '—'}
                              </td>
                            ))}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>

              <div style={{ display: 'grid', gap: 18 }}>
                <div style={{ ...panelStyle, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>Summary</div>
                  <div style={{ marginTop: 12, display: 'grid', gap: 10 }}>
                    {[
                      { label: 'Rows in sheet', value: active.rows.length },
                      { label: 'Filtered rows', value: rowsShown },
                      { label: 'Suspicious invoice', value: active.suspicious.length },
                      { label: 'Columns', value: active.headers.length },
                    ].map((item) => (
                      <div key={item.label} style={{ display: 'flex', justifyContent: 'space-between', gap: 12, padding: '10px 12px', borderRadius: 12, background: theme.card }}>
                        <span style={{ color: theme.muted }}>{item.label}</span>
                        <strong>{item.value}</strong>
                      </div>
                    ))}
                  </div>
                </div>

                <div style={{ ...panelStyle, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>Top numeric columns</div>
                  <div style={{ marginTop: 12, display: 'grid', gap: 10 }}>
                    {numericSummary.length ? numericSummary.map((item) => (
                      <div key={item.name} style={{ padding: 12, borderRadius: 12, background: theme.card, border: `1px solid ${theme.border}` }}>
                        <div style={{ color: theme.muted, fontSize: 12 }}>{item.name}</div>
                        <div style={{ fontWeight: 800, marginTop: 4 }}>{Number.isInteger(item.total) ? item.total.toLocaleString() : item.total.toFixed(2)}</div>
                        <div style={{ color: theme.muted, fontSize: 12 }}>{item.count} numeric cells</div>
                      </div>
                    )) : <div style={{ color: theme.muted }}>ไม่มีคอลัมน์ตัวเลข</div>}
                  </div>
                </div>

                <div style={{ ...panelStyle, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>Suspicious invoice list</div>
                  <div style={{ marginTop: 12, display: 'grid', gap: 10 }}>
                    {activeSuspicious.length ? activeSuspicious.map((item) => (
                      <div key={item.invoiceNo} style={{ padding: 12, borderRadius: 12, background: theme.card, border: `1px solid ${theme.border}` }}>
                        <div style={{ fontWeight: 800 }}>{item.invoiceNo}</div>
                        <div style={{ color: theme.muted, fontSize: 12, marginTop: 4 }}>{item.rows} rows · {item.skus} SKUs · {Number.isInteger(item.amount) ? item.amount.toLocaleString() : item.amount.toFixed(2)}</div>
                      </div>
                    )) : <div style={{ color: theme.muted }}>ไม่มีรายการเข้าข่าย</div>}
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {workbook && view === 'history' && (
          <div style={{ marginTop: 24, ...panelStyle, padding: 18 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
              <div>
                <div style={{ fontWeight: 800, fontSize: 18 }}>History</div>
                <div style={{ color: theme.muted, marginTop: 4 }}>บันทึกสรุปของไฟล์ที่เคยโหลดใน browser นี้</div>
              </div>
              <button
                onClick={() => setHistory([])}
                style={{ padding: '10px 14px', borderRadius: 12, border: `1px solid ${theme.border}`, background: theme.card, color: theme.text, cursor: 'pointer' }}
              >
                Clear history
              </button>
            </div>

            <div style={{ marginTop: 14, display: 'grid', gap: 10 }}>
              {history.length ? history.map((item) => (
                <div key={`${item.sourceName}-${item.parsedAt}`} style={{ padding: 14, borderRadius: 14, background: theme.card, border: `1px solid ${theme.border}` }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                    <div>
                      <div style={{ fontWeight: 800 }}>{item.sourceName}</div>
                      <div style={{ color: theme.muted, fontSize: 12 }}>{new Date(item.parsedAt).toLocaleString()}</div>
                    </div>
                    <div style={{ color: theme.muted, textAlign: 'right' }}>
                      {item.rows} rows · {item.sheets} sheets · {item.warnings} warnings · {item.inflationHits} suspicious
                    </div>
                  </div>
                </div>
              )) : <div style={{ color: theme.muted }}>ยังไม่มีประวัติ</div>}
            </div>
          </div>
        )}

        {view === 'settings' && (
          <div style={{ marginTop: 24, ...panelStyle, padding: 18 }}>
            <div style={{ fontWeight: 800, fontSize: 18 }}>Settings</div>
            <div style={{ display: 'grid', gap: 16, marginTop: 14, maxWidth: 720 }}>
              <label style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', padding: 14, borderRadius: 14, background: theme.card, border: `1px solid ${theme.border}` }}>
                <div>
                  <div style={{ fontWeight: 700 }}>Dark mode</div>
                  <div style={{ color: theme.muted, fontSize: 12 }}>สลับธีมของแอป</div>
                </div>
                <input type="checkbox" checked={darkMode} onChange={(e) => setDarkMode(e.target.checked)} />
              </label>

              <label style={{ display: 'grid', gap: 8, padding: 14, borderRadius: 14, background: theme.card, border: `1px solid ${theme.border}` }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}>
                  <div>
                    <div style={{ fontWeight: 700 }}>Rows per page</div>
                    <div style={{ color: theme.muted, fontSize: 12 }}>กำหนดจำนวนแถวที่แสดงในตาราง</div>
                  </div>
                  <strong>{rowsLimit}</strong>
                </div>
                <input type="range" min={25} max={500} step={25} value={rowsLimit} onChange={(e) => setRowsLimit(Number(e.target.value))} />
              </label>
            </div>
          </div>
        )}

        {!workbook && (
          <div style={{ marginTop: 24, ...panelStyle, padding: 18 }}>
            <div style={{ fontWeight: 800, fontSize: 18 }}>No workbook loaded</div>
            <div style={{ color: theme.muted, marginTop: 6 }}>อัปโหลดไฟล์ Excel เพื่อเริ่มอ่านข้อมูลจริง</div>
          </div>
        )}
      </div>
    </div>
  );
}
