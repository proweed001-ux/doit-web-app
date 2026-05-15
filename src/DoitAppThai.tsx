import JSZip from 'jszip';
import { useDeferredValue, useEffect, useMemo, useState } from 'react';
import { DEFAULT_DB, downloadJson, downloadText, safeParse, type DB, type Session } from './lib/dataset';
import { exportDatabase, parseDatabaseFile } from './lib/persistence';
import { filterRows, limitRows, summarizeHeaders } from './lib/runtimeOptimizations';
import {
  makeTemplate,
  parseSharedStrings,
  parseWorkbookRefs,
  safeText,
  splitCsv,
  colIndex,
  colName,
  type CanonicalField,
  type Cell,
  type Mapping,
  type Profile,
  type Row,
} from './lib/engine';
import {
  aggStoreTod,
  buildAnalysis,
  crossSheetRelations,
  numericSummary,
  type Relation,
  type RowAnalysis,
  type StoreStat,
  type TodStat,
} from './lib/analytics';

type View = 'dashboard' | 'sheets' | 'relations' | 'tod' | 'reports' | 'database' | 'settings';
type Theme = { shell: string; panel: string; card: string; border: string; text: string; muted: string; accent: string; danger: string; tableAlt: string };
type RawSheet = { name: string; headers: string[]; rows: Row[]; mapping: Mapping };
type SheetData = RawSheet & { anomalies: RowAnalysis[]; numericSummary: ReturnType<typeof numericSummary>; storeSummary: StoreStat[]; todSummary: TodStat[] };
type Workbook = { sourceName: string; parsedAt: string; sheets: SheetData[]; warnings: string[]; warningCount: number; rowCount: number; anomalyCount: number; relations: Relation[]; reportMarkdown: string };
type MappingKey = CanonicalField;

const AUTH_KEY = 'doit.auth.thai.v1';
const DB_KEY = 'doit.db.thai.v1';
const SETTINGS_KEY = 'doit.settings.thai.v1';

const fieldLabels: Record<MappingKey, string> = {
  invoiceNo: 'เลขที่บิล / เลขเอกสาร',
  invoiceAmount: 'ยอดเงินรวมของบิล',
  store: 'ชื่อร้าน / สาขา',
  date: 'วันที่',
  time: 'เวลา',
  sku: 'รหัสสินค้า',
  item: 'ชื่อสินค้า',
  qty: 'จำนวน',
  unitPrice: 'ราคาต่อชิ้น',
  lineTotal: 'ยอดรวมต่อรายการ',
  customer: 'ชื่อลูกค้า',
  category: 'กลุ่มสินค้า',
};

const viewLabels: Record<View, string> = {
  dashboard: 'หน้าหลัก',
  sheets: 'ดูตาราง',
  relations: 'เทียบข้อมูล',
  tod: 'ร้าน/ช่วงเวลา',
  reports: 'รายงาน',
  database: 'ฐานข้อมูล',
  settings: 'ตั้งค่า',
};

function theme(dark: boolean): Theme {
  return dark
    ? { shell: '#0f172a', panel: '#111827', card: '#0b1220', border: '#334155', text: '#e2e8f0', muted: '#94a3b8', accent: '#38bdf8', danger: '#fb7185', tableAlt: 'rgba(255,255,255,0.03)' }
    : { shell: '#f8fafc', panel: '#ffffff', card: '#f1f5f9', border: '#cbd5e1', text: '#0f172a', muted: '#475569', accent: '#0284c7', danger: '#e11d48', tableAlt: 'rgba(15,23,42,0.03)' };
}

function loadJson<T>(key: string, fallback: T): T {
  if (typeof window === 'undefined') return fallback;
  try { return safeParse<T>(window.localStorage.getItem(key), fallback); } catch { return fallback; }
}

function saveJson(key: string, value: unknown) {
  if (typeof window !== 'undefined') window.localStorage.setItem(key, JSON.stringify(value));
}

function formatNumber(value: number): string {
  if (!Number.isFinite(value)) return '0';
  return Number.isInteger(value) ? value.toLocaleString('th-TH') : value.toLocaleString('th-TH', { maximumFractionDigits: 2 });
}

function thaiReason(reason: string): string {
  const map: Record<string, string> = {
    'missing invoiceNo': 'ไม่พบเลขที่บิล',
    'missing invoiceAmount': 'ไม่พบยอดเงินรวม',
    'missing store': 'ไม่พบชื่อร้าน',
    'missing date': 'ไม่พบวันที่',
    'blank invoice': 'เลขที่บิลว่าง',
    'odd hour': 'เวลาผิดปกติ',
    'invalid amount': 'ยอดเงินไม่ถูกต้อง',
    'amount outlier': 'ยอดเงินสูง/ต่ำผิดปกติมาก',
    'amount deviation': 'ยอดเงินต่างจากปกติ',
    'duplicate invoice': 'เลขที่บิลซ้ำ',
    'qty*unit mismatch': 'จำนวนคูณราคาไม่ตรงกับยอดรวม',
  };
  return map[reason] ?? reason;
}

function thaiBucket(bucket: string): string {
  const map: Record<string, string> = {
    Unknown: 'ไม่ทราบเวลา',
    Night: 'กลางคืน',
    Morning: 'เช้า',
    Afternoon: 'บ่าย',
    Evening: 'เย็น',
    'Late Night': 'ดึก',
  };
  return map[bucket] ?? bucket;
}

function buildThaiReport(workbook: Workbook): string {
  const lines: string[] = [];
  lines.push('# รายงานตรวจข้อมูล DOIT');
  lines.push(`ไฟล์: ${workbook.sourceName}`);
  lines.push(`เวลาที่ตรวจ: ${new Date(workbook.parsedAt).toLocaleString('th-TH')}`);
  lines.push(`จำนวนชีต: ${formatNumber(workbook.sheets.length)}`);
  lines.push(`จำนวนแถวทั้งหมด: ${formatNumber(workbook.rowCount)}`);
  lines.push(`รายการที่ควรตรวจเพิ่ม: ${formatNumber(workbook.anomalyCount)}`);
  lines.push('');
  lines.push('## สรุปแต่ละชีต');
  workbook.sheets.forEach((sheet) => lines.push(`- ${sheet.name}: ${formatNumber(sheet.rows.length)} แถว, พบจุดน่าสงสัย ${formatNumber(sheet.anomalies.length)} รายการ`));
  lines.push('');
  lines.push('## รายการที่ควรตรวจเป็นพิเศษ');
  workbook.sheets.flatMap((sheet) => sheet.anomalies.slice(0, 5)).forEach((item) => {
    const reasons = item.reasons.map(thaiReason).join(', ') || 'ไม่มีหมายเหตุ';
    lines.push(`- ${item.sheet} :: ${item.invoiceNo || `แถวที่ ${item.rowIndex + 1}`} :: คะแนน ${item.score} :: ${reasons}`);
  });
  lines.push('');
  lines.push('## การเทียบข้อมูลข้ามชีต');
  if (!workbook.relations.length) lines.push('- ยังไม่พบข้อมูลที่เชื่อมกันชัดเจน');
  workbook.relations.forEach((relation) => lines.push(`- ${relation.leftSheet} เทียบกับ ${relation.rightSheet}: เจอข้อมูลซ้ำกัน ${formatNumber(relation.overlap)} จุด`));
  return lines.join('\n');
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
  const headers = (matrix[firstHeader] ?? []).map((v, i) => safeText(v) || `คอลัมน์ ${colName(i)}`);
  const rows = matrix.slice(firstHeader + 1).filter((r) => r?.some((v) => safeText(v).length > 0)).map((r) => {
    const obj: Row = {};
    headers.forEach((h, i) => { obj[h] = r?.[i] ?? ''; });
    return obj;
  });
  return { name, headers, rows };
}

function parseCsvSheet(text: string, name: string): { name: string; headers: string[]; rows: Row[] } {
  const lines = text.split(/\r?\n/).filter((line) => safeText(line).length > 0);
  const headers = splitCsv(lines[0] ?? '').map((h, i) => safeText(h) || `คอลัมน์ ${colName(i)}`);
  const rows = lines.slice(1).map((line) => {
    const cells = splitCsv(line);
    const row: Row = {};
    headers.forEach((h, i) => { row[h] = cells[i] ?? ''; });
    return row;
  });
  return { name, headers, rows };
}

function rebuildWorkbook(sourceName: string, parsedAt: string, sheets: RawSheet[], warnings: string[], profile: Profile): Workbook {
  const computed = sheets.map((sheet) => {
    const analysis = buildAnalysis(sheet.name, sheet.headers, sheet.rows, sheet.mapping, profile);
    const agg = aggStoreTod(sheet.headers, sheet.rows, sheet.mapping);
    return { ...sheet, anomalies: analysis.analyses, numericSummary: numericSummary(sheet.rows), storeSummary: agg.storeSummary, todSummary: agg.todSummary };
  });
  const relations = crossSheetRelations(computed.map((sheet) => ({ name: sheet.name, headers: sheet.headers, rows: sheet.rows })));
  const rowCount = computed.reduce((sum, sheet) => sum + sheet.rows.length, 0);
  const anomalyCount = computed.reduce((sum, sheet) => sum + sheet.anomalies.length, 0);
  const workbook: Workbook = { sourceName, parsedAt, sheets: computed, warnings, warningCount: warnings.length, rowCount, anomalyCount, relations, reportMarkdown: '' };
  workbook.reportMarkdown = buildThaiReport(workbook);
  return workbook;
}

async function parseWorkbookFile(file: File, db: DB, profile: Profile): Promise<Workbook> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const workbookXml = await zip.file('xl/workbook.xml')?.async('text');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('text');
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('text');
  if (!workbookXml || !relsXml) throw new Error('ไฟล์นี้ไม่ใช่ไฟล์ Excel มาตรฐาน หรือไฟล์อาจเสีย');
  const sharedStrings = parseSharedStrings(sharedStringsXml ?? '');
  const refs = parseWorkbookRefs(workbookXml, relsXml);
  const warnings: string[] = [];
  const sheets: RawSheet[] = [];
  for (const ref of refs) {
    const xml = await zip.file(ref.path)?.async('text');
    if (!xml) { warnings.push(`อ่านชีต ${ref.name} ไม่ได้`); continue; }
    const parsed = parseWorksheet(xml, sharedStrings, ref.name);
    const mapping = { ...makeTemplate(parsed.headers), ...(db.templates[ref.name] ?? {}) };
    sheets.push({ ...parsed, mapping });
    db.templates[ref.name] = mapping;
  }
  if (!sheets.length) throw new Error('ไม่พบตารางข้อมูลที่อ่านได้ในไฟล์นี้');
  return rebuildWorkbook(file.name, new Date().toISOString(), sheets, warnings, profile);
}

async function parseCsvFile(file: File, profile: Profile): Promise<Workbook> {
  const parsed = parseCsvSheet(await file.text(), file.name);
  const mapping = makeTemplate(parsed.headers);
  return rebuildWorkbook(file.name, new Date().toISOString(), [{ ...parsed, mapping }], ['อ่านแบบไฟล์ CSV'], profile);
}

export default function DoitAppThai() {
  const settings = loadJson<{ dark: boolean; rowsLimit: number; profile: Profile }>(SETTINGS_KEY, { dark: true, rowsLimit: 100, profile: 'balanced' });
  const [session, setSession] = useState<Session | null>(loadJson<Session | null>(AUTH_KEY, null));
  const [db, setDB] = useState<DB>(loadJson<DB>(DB_KEY, DEFAULT_DB));
  const [dark, setDark] = useState(settings.dark);
  const [rowsLimit, setRowsLimit] = useState(settings.rowsLimit);
  const [profile, setProfile] = useState<Profile>(settings.profile);
  const [view, setView] = useState<View>('dashboard');
  const [status, setStatus] = useState('พร้อมใช้งาน');
  const [error, setError] = useState('');
  const [query, setQuery] = useState('');
  const deferredQuery = useDeferredValue(query);
  const [workbook, setWorkbook] = useState<Workbook | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);

  const t = theme(dark);
  const active = workbook?.sheets[activeSheet] ?? null;
  const filteredRows = useMemo(() => (active ? filterRows(active.rows, deferredQuery) : []), [active, deferredQuery]);
  const visibleRows = useMemo(() => limitRows(filteredRows, rowsLimit), [filteredRows, rowsLimit]);
  const workbookSnapshots = useMemo(() => {
    const sheets = workbook?.sheets ?? [];
    return {
      anomalies: sheets.flatMap((sheet) => sheet.anomalies),
      stores: sheets.flatMap((sheet) => sheet.storeSummary.map((item) => ({ ...item, sheet: sheet.name }))),
      tod: sheets.flatMap((sheet) => sheet.todSummary.map((item) => ({ ...item, sheet: sheet.name }))),
      report: workbook?.reportMarkdown ?? '',
      relations: workbook?.relations ?? [],
    };
  }, [workbook]);

  useEffect(() => {
    saveJson(AUTH_KEY, session);
    saveJson(DB_KEY, db);
    saveJson(SETTINGS_KEY, { dark, rowsLimit, profile });
    if (typeof document !== 'undefined') {
      document.documentElement.lang = 'th';
      document.documentElement.style.colorScheme = dark ? 'dark' : 'light';
      document.body.style.margin = '0';
      document.body.style.background = t.shell;
    }
  }, [session, db, dark, rowsLimit, profile, t.shell]);

  useEffect(() => {
    if (!workbook) return;
    setWorkbook((current) => current ? rebuildWorkbook(current.sourceName, current.parsedAt, current.sheets.map((sheet) => ({ name: sheet.name, headers: sheet.headers, rows: sheet.rows, mapping: sheet.mapping })), current.warnings, profile) : current);
  }, [profile]);

  function login() {
    const name = (document.getElementById('name') as HTMLInputElement | null)?.value ?? '';
    const org = (document.getElementById('org') as HTMLInputElement | null)?.value ?? '';
    const next: Session = { name: safeText(name) || 'ผู้ใช้งาน', org: safeText(org) || 'ทีมงาน', token: crypto.randomUUID(), createdAt: new Date().toISOString() };
    setSession(next);
    setDB((prev) => ({ ...prev, sessions: [next, ...prev.sessions].slice(0, 20) }));
    setStatus(`เข้าใช้งานแล้ว: ${next.name}`);
  }

  function logout() {
    setSession(null);
    setWorkbook(null);
    setStatus('ออกจากระบบแล้ว');
  }

  function persistWorkbook(next: Workbook) {
    setWorkbook(next);
    setDB((prev) => ({
      ...prev,
      workbooks: [{ name: next.sourceName, parsedAt: next.parsedAt, sheets: next.sheets.length, rows: next.rowCount, anomalies: next.anomalyCount }, ...prev.workbooks].slice(0, 20),
      reports: [{ title: next.sourceName, generatedAt: next.parsedAt, markdown: next.reportMarkdown, source: next.sourceName }, ...prev.reports].slice(0, 20),
    }));
  }

  async function handleLoadFile(file: File) {
    if (!session) { setError('กรุณากดเข้าใช้งานก่อน'); return; }
    setError('');
    setStatus(`กำลังอ่านไฟล์ ${file.name}`);
    try {
      const wb = /\.(xlsx|xlsm|zip)$/i.test(file.name) ? await parseWorkbookFile(file, { ...db }, profile) : await parseCsvFile(file, profile);
      persistWorkbook(wb);
      setActiveSheet(0);
      setView('dashboard');
      setStatus(`อ่านไฟล์สำเร็จ: พบ ${formatNumber(wb.sheets.length)} ชีต`);
    } catch (e) {
      setWorkbook(null);
      setError((e as Error).message || 'อ่านไฟล์ไม่สำเร็จ กรุณาตรวจไฟล์อีกครั้ง');
      setStatus('เกิดข้อผิดพลาด');
    }
  }

  async function handleImportDatabase(file: File) {
    if (!session) { setError('กรุณากดเข้าใช้งานก่อน'); return; }
    try {
      setDB(parseDatabaseFile(await file.text()));
      setStatus('นำเข้าฐานข้อมูลสำเร็จ');
    } catch (e) {
      setError((e as Error).message || 'นำเข้าฐานข้อมูลไม่สำเร็จ');
    }
  }

  function handleExportDatabase() { downloadText('ฐานข้อมูล-doit.json', exportDatabase(db), 'application/json;charset=utf-8'); }
  function exportReport() { if (workbook) downloadText(`${workbook.sourceName.replace(/\.[^.]+$/, '')}-รายงาน.md`, workbook.reportMarkdown, 'text/markdown;charset=utf-8'); }
  function exportJson() { if (workbook) downloadJson(`${workbook.sourceName.replace(/\.[^.]+$/, '')}-ข้อมูล.json`, workbook); }
  function saveTemplate() { if (!active) return; setDB((prev) => ({ ...prev, templates: { ...prev.templates, [active.name]: active.mapping } })); setStatus(`บันทึกการจับคู่หัวตารางของ ${active.name} แล้ว`); }

  function updateMapping(field: MappingKey, header: string) {
    if (!active || !workbook) return;
    const nextSheets = workbook.sheets.map((sheet, idx) => idx === activeSheet ? { ...sheet, mapping: { ...sheet.mapping, [field]: header || undefined } } : sheet);
    const next = rebuildWorkbook(workbook.sourceName, workbook.parsedAt, nextSheets.map((sheet) => ({ name: sheet.name, headers: sheet.headers, rows: sheet.rows, mapping: sheet.mapping })), workbook.warnings, profile);
    persistWorkbook(next);
    setDB((prev) => ({ ...prev, templates: { ...prev.templates, [active.name]: { ...active.mapping, [field]: header || undefined } } }));
  }

  const buttonStyle = (id: View) => ({ padding: '12px 16px', borderRadius: 999, border: `1px solid ${view === id ? t.accent : t.border}`, background: view === id ? `${t.accent}22` : t.panel, color: t.text, fontWeight: 800 as const, cursor: 'pointer', fontSize: 16 });
  const shell = { minHeight: '100vh', background: t.shell, color: t.text, fontFamily: 'Tahoma, Noto Sans Thai, system-ui, sans-serif', fontSize: 16 } as const;
  const panel = { background: t.panel, border: `1px solid ${t.border}`, borderRadius: 20, boxShadow: dark ? '0 20px 60px rgba(0,0,0,0.18)' : '0 18px 48px rgba(15,23,42,0.06)' } as const;
  const card = { background: t.card, border: `1px solid ${t.border}`, borderRadius: 16 } as const;

  return (
    <div style={shell}>
      <div style={{ maxWidth: 1480, margin: '0 auto', padding: 20 }}>
        {!session ? (
          <div style={{ ...panel, maxWidth: 720, margin: '48px auto', padding: 24 }}>
            <h1 style={{ margin: 0, fontSize: 'clamp(2rem, 4vw, 3rem)' }}>ระบบตรวจข้อมูล DOIT</h1>
            <p style={{ color: t.muted, fontSize: 18, lineHeight: 1.6 }}>กดเข้าใช้งานก่อน แล้วค่อยเลือกไฟล์ Excel เพื่อให้ระบบช่วยสรุปยอด ตรวจข้อมูลซ้ำ และชี้จุดที่ควรดูเพิ่ม</p>
            <div style={{ display: 'grid', gap: 12, marginTop: 16 }}>
              <input id="name" placeholder="ชื่อผู้ใช้ เช่น พี่ฐา" style={{ padding: 14, borderRadius: 12, border: `1px solid ${t.border}`, background: t.card, color: t.text, fontSize: 18 }} />
              <input id="org" placeholder="ชื่อทีม/หน่วยงาน เช่น ฝ่ายขาย" style={{ padding: 14, borderRadius: 12, border: `1px solid ${t.border}`, background: t.card, color: t.text, fontSize: 18 }} />
              <button onClick={login} style={{ padding: 14, borderRadius: 12, border: 'none', background: t.accent, color: '#082f49', fontWeight: 900, cursor: 'pointer', fontSize: 18 }}>เข้าใช้งาน</button>
            </div>
          </div>
        ) : (
          <>
            <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap', alignItems: 'flex-end' }}>
              <div>
                <div style={{ fontSize: 14, color: t.muted, letterSpacing: 0.4 }}>ระบบ DOIT</div>
                <h1 style={{ margin: '8px 0 6px', fontSize: 'clamp(2rem, 3vw, 3.4rem)', lineHeight: 1.05 }}>หน้าควบคุมข้อมูล</h1>
                <div style={{ color: t.muted, fontSize: 17 }}>ผู้ใช้: <strong>{session.name}</strong> · ทีม: {session.org}</div>
              </div>
              <div style={{ ...panel, minWidth: 280, padding: 16 }}>
                <div style={{ fontSize: 13, color: t.muted }}>สถานะ</div>
                <div style={{ fontWeight: 900, fontSize: 18 }}>{status}</div>
                {workbook && <div style={{ marginTop: 6, fontSize: 13, color: t.muted }}>{workbook.sourceName} · {new Date(workbook.parsedAt).toLocaleString('th-TH')}</div>}
              </div>
            </div>

            <div style={{ marginTop: 18, ...panel, padding: 18 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', gap: 16, flexWrap: 'wrap', alignItems: 'center' }}>
                <div>
                  <div style={{ fontWeight: 900, fontSize: 20 }}>เลือกไฟล์ข้อมูล</div>
                  <div style={{ color: t.muted, fontSize: 17 }}>รองรับไฟล์ Excel, CSV และไฟล์ฐานข้อมูลที่เคยบันทึกไว้</div>
                </div>
                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                  <label style={{ cursor: 'pointer', padding: '14px 18px', borderRadius: 14, background: t.accent, color: '#082f49', fontWeight: 900, fontSize: 18 }}>
                    เลือกไฟล์ Excel
                    <input type="file" accept=".xlsx,.xlsm,.zip,.csv,.txt" style={{ display: 'none' }} onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleLoadFile(f); }} />
                  </label>
                  <label style={{ cursor: 'pointer', padding: '14px 18px', borderRadius: 14, background: t.panel, color: t.text, border: `1px solid ${t.border}`, fontWeight: 900, fontSize: 18 }}>
                    นำเข้าฐานข้อมูล
                    <input type="file" accept="application/json,.json" style={{ display: 'none' }} onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleImportDatabase(f); }} />
                  </label>
                </div>
              </div>
              {error && <div style={{ marginTop: 12, padding: 14, borderRadius: 12, background: `${t.danger}18`, color: t.danger, fontWeight: 800 }}>{error}</div>}
            </div>

            <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 16 }}>
              {(['dashboard', 'sheets', 'relations', 'tod', 'reports', 'database', 'settings'] as View[]).map((v) => <button key={v} style={buttonStyle(v)} onClick={() => setView(v)}>{viewLabels[v]}</button>)}
              <button style={buttonStyle('reports')} onClick={exportReport} disabled={!workbook}>โหลดรายงาน</button>
              <button style={buttonStyle('reports')} onClick={exportJson} disabled={!workbook}>โหลดข้อมูล</button>
              <button style={buttonStyle('database')} onClick={logout}>ออกจากระบบ</button>
            </div>

            {!workbook && <div style={{ ...panel, marginTop: 22, padding: 24, color: t.muted, fontSize: 20 }}>ยังไม่มีไฟล์ข้อมูล ให้กดปุ่ม “เลือกไฟล์ Excel” ด้านบนก่อน</div>}

            {workbook && view === 'dashboard' && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 16 }}>
                  {[
                    { label: 'จำนวนชีต', value: workbook.sheets.length },
                    { label: 'จำนวนแถว', value: workbook.rowCount },
                    { label: 'ควรตรวจเพิ่ม', value: workbook.anomalyCount },
                    { label: 'ข้อมูลเชื่อมกัน', value: workbook.relations.length },
                  ].map((item) => <div key={item.label} style={{ ...card, padding: 18 }}><div style={{ color: t.muted, fontSize: 15 }}>{item.label}</div><div style={{ fontSize: 34, fontWeight: 900, marginTop: 4 }}>{formatNumber(item.value)}</div></div>)}
                </div>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ fontWeight: 900, fontSize: 20 }}>รายการที่ควรตรวจเป็นพิเศษ</div>
                  <div style={{ color: t.muted, marginTop: 4 }}>คะแนนสูงแปลว่าควรเปิดดู แต่ไม่ได้แปลว่าผิดเสมอไป</div>
                  <div style={{ marginTop: 12, display: 'grid', gap: 10 }}>
                    {workbookSnapshots.anomalies.slice(0, 12).map((item) => <div key={`${item.sheet}-${item.rowIndex}`} style={{ padding: 14, borderRadius: 12, background: card.background, border: `1px solid ${t.border}` }}><div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}><strong>{item.sheet} · {item.invoiceNo || `แถวที่ ${item.rowIndex + 1}`}</strong><span style={{ color: t.danger, fontWeight: 900 }}>คะแนน {item.score}</span></div><div style={{ color: t.muted, marginTop: 6 }}>{item.reasons.map(thaiReason).join(' · ') || 'ปกติ'}</div></div>)}
                    {!workbookSnapshots.anomalies.length && <div style={{ color: t.muted }}>ยังไม่พบรายการน่าสงสัย</div>}
                  </div>
                </div>
              </div>
            )}

            {workbook && view === 'sheets' && active && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>{workbook.sheets.map((sheet, idx) => <button key={sheet.name + idx} onClick={() => setActiveSheet(idx)} style={{ ...buttonStyle('sheets'), background: idx === activeSheet ? `${t.accent}22` : t.panel }}>{sheet.name}</button>)}</div>
                <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1.6fr) minmax(280px, 0.9fr)', gap: 18 }}>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ ...panel, padding: 18 }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                        <div><div style={{ fontSize: 14, color: t.muted }}>ตารางที่กำลังดู</div><div style={{ fontSize: 22, fontWeight: 900 }}>{active.name}</div></div>
                        <div style={{ color: t.muted }}>{formatNumber(active.rows.length)} แถว · {formatNumber(active.headers.length)} คอลัมน์ · {summarizeHeaders(active.headers)}</div>
                      </div>
                      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 14 }}>
                        <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="ค้นหาชื่อร้าน เลขบิล หรือสินค้า..." style={{ flex: '1 1 220px', padding: 14, borderRadius: 12, border: `1px solid ${t.border}`, background: t.card, color: t.text, fontSize: 18 }} />
                        <button onClick={saveTemplate} style={buttonStyle('sheets')}>จำหัวตารางนี้</button>
                      </div>
                    </div>
                    <div style={{ marginTop: 16, overflowX: 'auto', borderRadius: 18, border: `1px solid ${t.border}`, background: t.panel }}>
                      <table style={{ width: '100%', minWidth: 900, borderCollapse: 'collapse' }}>
                        <thead><tr>{active.headers.map((h) => <th key={h} style={{ position: 'sticky', top: 0, background: t.panel, textAlign: 'left', padding: 12, borderBottom: `1px solid ${t.border}`, color: t.muted }}>{h}</th>)}</tr></thead>
                        <tbody>{visibleRows.map((row, i) => <tr key={i} style={{ background: i % 2 ? t.tableAlt : 'transparent' }}>{active.headers.map((h) => <td key={h} style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{safeText(row[h]) || '—'}</td>)}</tr>)}</tbody>
                      </table>
                    </div>
                  </div>
                  <div style={{ display: 'grid', gap: 18 }}>
                    <div style={{ ...panel, padding: 18 }}>
                      <div style={{ fontWeight: 900, fontSize: 20 }}>จับคู่หัวตาราง</div>
                      <div style={{ color: t.muted, marginTop: 4 }}>เลือกว่าคอลัมน์ไหนคืออะไร ถ้าระบบจับผิดให้แก้ตรงนี้</div>
                      <div style={{ display: 'grid', gap: 10, marginTop: 12 }}>
                        {(Object.keys(makeTemplate(active.headers)) as MappingKey[]).map((field) => (
                          <label key={field} style={{ display: 'grid', gap: 6 }}>
                            <span style={{ color: t.muted, fontSize: 14 }}>{fieldLabels[field]}</span>
                            <select value={active.mapping[field] ?? ''} onChange={(e) => updateMapping(field, e.target.value)} style={{ padding: 12, borderRadius: 12, border: `1px solid ${t.border}`, background: t.card, color: t.text, fontSize: 16 }}>
                              <option value="">ไม่เลือก</option>
                              {active.headers.map((h) => <option key={h} value={h}>{h}</option>)}
                            </select>
                          </label>
                        ))}
                      </div>
                    </div>
                    <div style={{ ...panel, padding: 18 }}>
                      <div style={{ fontWeight: 900, fontSize: 20 }}>สรุปตารางนี้</div>
                      <div style={{ display: 'grid', gap: 8, marginTop: 12 }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>จำนวนแถว</span><strong>{formatNumber(active.rows.length)}</strong></div>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>ควรตรวจเพิ่ม</span><strong>{formatNumber(active.anomalies.length)}</strong></div>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>ร้านที่สรุปได้</span><strong>{formatNumber(active.storeSummary.length)}</strong></div>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>ช่วงเวลาที่พบ</span><strong>{formatNumber(active.todSummary.length)}</strong></div>
                      </div>
                    </div>
                    <div style={{ ...panel, padding: 18 }}>
                      <div style={{ fontWeight: 900, fontSize: 20 }}>คอลัมน์ตัวเลขสำคัญ</div>
                      <div style={{ display: 'grid', gap: 10, marginTop: 12 }}>{active.numericSummary.slice(0, 8).map((n) => <div key={n.name} style={{ padding: 12, borderRadius: 12, background: t.card, border: `1px solid ${t.border}` }}><div style={{ color: t.muted, fontSize: 14 }}>{n.name}</div><strong>{formatNumber(n.total)}</strong><div style={{ color: t.muted, fontSize: 13 }}>เจอข้อมูลตัวเลข {formatNumber(n.count)} ช่อง</div></div>)}</div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {workbook && view === 'relations' && <div style={{ marginTop: 22, ...panel, padding: 18 }}><div style={{ fontWeight: 900, fontSize: 20 }}>เทียบข้อมูลข้ามตาราง</div><div style={{ color: t.muted, marginTop: 4 }}>ใช้ดูว่าชีตไหนมีข้อมูลซ้ำหรือเกี่ยวข้องกัน</div><div style={{ marginTop: 12, display: 'grid', gap: 10 }}>{workbookSnapshots.relations.length ? workbookSnapshots.relations.map((relation, idx) => <div key={idx} style={{ padding: 14, borderRadius: 12, background: card.background, border: `1px solid ${t.border}` }}><strong>{relation.leftSheet} เทียบกับ {relation.rightSheet}</strong><div style={{ color: t.muted, fontSize: 14 }}>พบข้อมูลตรงกัน {formatNumber(relation.overlap)} จุด</div></div>) : <div style={{ color: t.muted }}>ยังไม่พบข้อมูลที่เชื่อมกันชัดเจน</div>}</div></div>}

            {workbook && view === 'tod' && <div style={{ marginTop: 22, display: 'grid', gap: 18 }}><div style={{ ...panel, padding: 18 }}><div style={{ fontWeight: 900, fontSize: 20 }}>สรุปตามร้าน</div><div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 16, marginTop: 12 }}>{workbookSnapshots.stores.slice(0, 8).map((store, idx) => <div key={idx} style={{ ...card, padding: 14 }}><div style={{ fontWeight: 900 }}>{store.store}</div><div style={{ color: t.muted, fontSize: 13 }}>{store.sheet}</div><div style={{ marginTop: 8 }}>จำนวนรายการ {formatNumber(store.count)}</div><div>ยอดเงิน {formatNumber(store.amount)}</div><div style={{ color: t.danger }}>ควรตรวจ {formatNumber(store.suspicious)}</div></div>)}</div></div><div style={{ ...panel, padding: 18 }}><div style={{ fontWeight: 900, fontSize: 20 }}>สรุปตามช่วงเวลา</div><div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12, marginTop: 12 }}>{workbookSnapshots.tod.map((item, idx) => <div key={idx} style={{ ...card, padding: 14 }}><div style={{ fontWeight: 900 }}>{thaiBucket(item.bucket)}</div><div style={{ color: t.muted, fontSize: 13 }}>{item.sheet}</div><div>จำนวนรายการ {formatNumber(item.count)}</div><div>ยอดเงิน {formatNumber(item.amount)}</div></div>)}</div></div></div>}

            {workbook && view === 'reports' && <div style={{ marginTop: 22, display: 'grid', gap: 18 }}><div style={{ ...panel, padding: 18 }}><div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}><div><div style={{ fontWeight: 900, fontSize: 20 }}>รายงานสรุป</div><div style={{ color: t.muted, marginTop: 4 }}>อ่านง่าย ใช้ส่งต่อหรือเก็บไว้ตรวจย้อนหลังได้</div></div><div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}><button onClick={exportReport} style={buttonStyle('reports')}>โหลดรายงาน</button><button onClick={exportJson} style={buttonStyle('reports')}>โหลดข้อมูล</button></div></div><pre style={{ whiteSpace: 'pre-wrap', marginTop: 16, padding: 16, borderRadius: 14, background: t.card, border: `1px solid ${t.border}`, overflowX: 'auto', lineHeight: 1.6, fontFamily: 'Tahoma, Noto Sans Thai, system-ui, sans-serif' }}>{workbookSnapshots.report}</pre></div></div>}

            {view === 'database' && <div style={{ marginTop: 22, display: 'grid', gap: 18 }}><div style={{ ...panel, padding: 18 }}><div style={{ fontWeight: 900, fontSize: 20 }}>ฐานข้อมูลในเครื่องนี้</div><div style={{ color: t.muted, marginTop: 4 }}>ข้อมูลถูกเก็บไว้ในเครื่อง/เบราว์เซอร์นี้ ถ้าเปลี่ยนเครื่องต้องส่งออกไฟล์ฐานข้อมูลก่อน</div><div style={{ display: 'grid', gap: 10, marginTop: 12 }}><div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>ผู้ใช้ที่เคยเข้า</span><strong>{formatNumber(db.sessions.length)}</strong></div><div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>ไฟล์ที่เคยอ่าน</span><strong>{formatNumber(db.workbooks.length)}</strong></div><div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>รายงานที่เก็บไว้</span><strong>{formatNumber(db.reports.length)}</strong></div><div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>หัวตารางที่จำไว้</span><strong>{formatNumber(Object.keys(db.templates).length)}</strong></div></div><div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 14 }}><button onClick={handleExportDatabase} style={buttonStyle('database')}>ส่งออกฐานข้อมูล</button><label style={{ cursor: 'pointer', padding: '12px 16px', borderRadius: 12, border: `1px solid ${t.border}`, background: t.panel, color: t.text, fontWeight: 800 }}>นำเข้าฐานข้อมูล<input type="file" accept="application/json,.json" style={{ display: 'none' }} onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleImportDatabase(f); }} /></label><button onClick={() => { setDB(DEFAULT_DB); setStatus('ล้างฐานข้อมูลแล้ว'); }} style={buttonStyle('database')}>ล้างฐานข้อมูล</button></div></div><div style={{ ...panel, padding: 18 }}><div style={{ fontWeight: 900, fontSize: 20 }}>ผู้ใช้งาน</div><div style={{ marginTop: 10 }}>ชื่อ: <strong>{session.name}</strong></div><div>ทีม: <strong>{session.org}</strong></div></div></div>}

            {view === 'settings' && <div style={{ marginTop: 22, ...panel, padding: 18 }}><div style={{ fontWeight: 900, fontSize: 20 }}>ตั้งค่าระบบ</div><div style={{ display: 'grid', gap: 16, marginTop: 14, maxWidth: 760 }}><label style={{ display: 'grid', gap: 8, padding: 14, borderRadius: 14, background: t.card, border: `1px solid ${t.border}` }}><div style={{ display: 'flex', justifyContent: 'space-between' }}><span>โหมดกลางคืน</span><strong>{dark ? 'เปิด' : 'ปิด'}</strong></div><input type="checkbox" checked={dark} onChange={(e) => setDark(e.target.checked)} /></label><label style={{ display: 'grid', gap: 8, padding: 14, borderRadius: 14, background: t.card, border: `1px solid ${t.border}` }}><div style={{ display: 'flex', justifyContent: 'space-between' }}><span>จำนวนแถวที่แสดง</span><strong>{formatNumber(rowsLimit)}</strong></div><input type="range" min={25} max={500} step={25} value={rowsLimit} onChange={(e) => setRowsLimit(Number(e.target.value))} /></label><label style={{ display: 'grid', gap: 8, padding: 14, borderRadius: 14, background: t.card, border: `1px solid ${t.border}` }}><div style={{ display: 'flex', justifyContent: 'space-between' }}><span>ระดับการตรวจ</span><strong>{profile === 'balanced' ? 'ปกติ' : profile === 'strict' ? 'เข้มงวด' : 'ละเอียดมาก'}</strong></div><select value={profile} onChange={(e) => setProfile(e.target.value as Profile)} style={{ padding: 12, borderRadius: 12, border: `1px solid ${t.border}`, background: t.panel, color: t.text }}><option value="balanced">ปกติ</option><option value="strict">เข้มงวด</option><option value="forensic">ละเอียดมาก</option></select></label></div></div>}
          </>
        )}
      </div>
    </div>
  );
}
