import JSZip from 'jszip';
import { useEffect, useMemo, useState } from 'react';
import { DEFAULT_DB, downloadJson, downloadText, safeParse, type DB, type Session } from './lib/dataset';
import { exportDatabase, parseDatabaseFile } from './lib/persistence';
import {
  makeTemplate,
  normalize,
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
  buildReport,
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

const AUTH_KEY = 'doit.auth.v2';
const DB_KEY = 'doit.db.v2';
const SETTINGS_KEY = 'doit.settings.v2';

function theme(dark: boolean): Theme {
  return dark
    ? { shell: '#0f172a', panel: '#111827', card: '#0b1220', border: '#334155', text: '#e2e8f0', muted: '#94a3b8', accent: '#38bdf8', danger: '#fb7185', tableAlt: 'rgba(255,255,255,0.03)' }
    : { shell: '#f8fafc', panel: '#ffffff', card: '#f1f5f9', border: '#cbd5e1', text: '#0f172a', muted: '#475569', accent: '#0284c7', danger: '#e11d48', tableAlt: 'rgba(15,23,42,0.03)' };
}

function loadJson<T>(key: string, fallback: T): T {
  if (typeof window === 'undefined') return fallback;
  try {
    return safeParse<T>(window.localStorage.getItem(key), fallback);
  } catch {
    return fallback;
  }
}

function saveJson(key: string, value: unknown) {
  if (typeof window === 'undefined') return;
  window.localStorage.setItem(key, JSON.stringify(value));
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

function computeSheet(sheet: RawSheet, profile: Profile): SheetData {
  const analysis = buildAnalysis(sheet.name, sheet.headers, sheet.rows, sheet.mapping, profile);
  const agg = aggStoreTod(sheet.headers, sheet.rows, sheet.mapping);
  return { ...sheet, anomalies: analysis.analyses, numericSummary: numericSummary(sheet.rows), storeSummary: agg.storeSummary, todSummary: agg.todSummary };
}

function computeWorkbook(sourceName: string, parsedAt: string, sheets: RawSheet[], warnings: string[], profile: Profile): Workbook {
  const computed = sheets.map((s) => computeSheet(s, profile));
  const relations = crossSheetRelations(computed.map((s) => ({ name: s.name, headers: s.headers, rows: s.rows })));
  const rowCount = computed.reduce((sum, s) => sum + s.rows.length, 0);
  const anomalyCount = computed.reduce((sum, s) => sum + s.anomalies.length, 0);
  const workbook: Workbook = { sourceName, parsedAt, sheets: computed, warnings, warningCount: warnings.length, rowCount, anomalyCount, relations, reportMarkdown: '' };
  workbook.reportMarkdown = buildReport(workbook);
  return workbook;
}

async function parseWorkbookFile(file: File, db: DB, profile: Profile): Promise<Workbook> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const workbookXml = await zip.file('xl/workbook.xml')?.async('text');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('text');
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('text');
  if (!workbookXml || !relsXml) throw new Error('ไฟล์ไม่ใช่ workbook มาตรฐาน');
  const sharedStrings = parseSharedStrings(sharedStringsXml ?? '');
  const refs = parseWorkbookRefs(workbookXml, relsXml);
  const warnings: string[] = [];
  const sheets: RawSheet[] = [];
  for (const ref of refs) {
    const xml = await zip.file(ref.path)?.async('text');
    if (!xml) { warnings.push(`Missing sheet data: ${ref.name}`); continue; }
    const parsed = parseWorksheet(xml, sharedStrings, ref.name);
    const mapping = { ...makeTemplate(parsed.headers), ...(db.templates[ref.name] ?? {}) };
    sheets.push({ ...parsed, mapping });
    db.templates[ref.name] = mapping;
  }
  if (!sheets.length) throw new Error('ไม่พบข้อมูลที่ใช้งานได้ในไฟล์');
  return computeWorkbook(file.name, new Date().toISOString(), sheets, warnings, profile);
}

async function parseCsvFile(file: File, profile: Profile): Promise<Workbook> {
  const parsed = parseCsvSheet(await file.text(), file.name);
  const mapping = makeTemplate(parsed.headers);
  return computeWorkbook(file.name, new Date().toISOString(), [{ ...parsed, mapping }], ['CSV fallback mode'], profile);
}

function readLocalDB(): DB { return loadJson<DB>(DB_KEY, DEFAULT_DB); }
function writeLocalDB(db: DB) { saveJson(DB_KEY, db); }

export default function DoitAppDb() {
  const settings = loadJson<{ dark: boolean; rowsLimit: number; profile: Profile }>(SETTINGS_KEY, { dark: true, rowsLimit: 100, profile: 'balanced' });
  const [session, setSession] = useState<Session | null>(loadJson<Session | null>(AUTH_KEY, null));
  const [db, setDB] = useState<DB>(readLocalDB());
  const [dark, setDark] = useState(settings.dark);
  const [rowsLimit, setRowsLimit] = useState(settings.rowsLimit);
  const [profile, setProfile] = useState<Profile>(settings.profile);
  const [view, setView] = useState<View>('dashboard');
  const [status, setStatus] = useState('พร้อมใช้งาน');
  const [error, setError] = useState('');
  const [query, setQuery] = useState('');
  const [workbook, setWorkbook] = useState<Workbook | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);

  const t = theme(dark);
  const active = workbook?.sheets[activeSheet] ?? null;
  const filteredRows = useMemo(() => {
    if (!active) return [] as Row[];
    if (!query.trim()) return active.rows;
    const q = normalize(query);
    return active.rows.filter((row) => Object.entries(row).some(([k, v]) => normalize(k).includes(q) || normalize(v).includes(q)));
  }, [active, query]);
  const visibleRows = filteredRows.slice(0, rowsLimit);
  const allAnomalies = workbook?.sheets.flatMap((s) => s.anomalies) ?? [];
  const allStores = workbook?.sheets.flatMap((s) => s.storeSummary.map((x) => ({ ...x, sheet: s.name }))) ?? [];
  const allTod = workbook?.sheets.flatMap((s) => s.todSummary.map((x) => ({ ...x, sheet: s.name }))) ?? [];

  useEffect(() => {
    saveJson(AUTH_KEY, session);
    writeLocalDB(db);
    saveJson(SETTINGS_KEY, { dark, rowsLimit, profile });
    if (typeof document !== 'undefined') {
      document.documentElement.style.colorScheme = dark ? 'dark' : 'light';
      document.body.style.margin = '0';
      document.body.style.background = t.shell;
    }
  }, [session, db, dark, rowsLimit, profile, t.shell]);

  useEffect(() => {
    if (!workbook) return;
    setWorkbook((current) => current ? computeWorkbook(current.sourceName, current.parsedAt, current.sheets.map((s) => ({ name: s.name, headers: s.headers, rows: s.rows, mapping: s.mapping })), current.warnings, profile) : current);
  }, [profile]);

  function login() {
    const name = (document.getElementById('name') as HTMLInputElement | null)?.value ?? '';
    const org = (document.getElementById('org') as HTMLInputElement | null)?.value ?? '';
    const next: Session = { name: safeText(name) || 'User', org: safeText(org) || 'Workspace', token: crypto.randomUUID(), createdAt: new Date().toISOString() };
    setSession(next);
    setDB((prev) => ({ ...prev, sessions: [next, ...prev.sessions].slice(0, 20) }));
    setStatus(`เข้าสู่ระบบแล้ว: ${next.name}`);
  }

  function logout() { setSession(null); setWorkbook(null); setStatus('ออกจากระบบ'); }

  function persistWorkbook(next: Workbook) {
    setWorkbook(next);
    setDB((prev) => ({
      ...prev,
      workbooks: [{ name: next.sourceName, parsedAt: next.parsedAt, sheets: next.sheets.length, rows: next.rowCount, anomalies: next.anomalyCount }, ...prev.workbooks].slice(0, 20),
      reports: [{ title: next.sourceName, generatedAt: next.parsedAt, markdown: next.reportMarkdown, source: next.sourceName }, ...prev.reports].slice(0, 20),
    }));
  }

  async function handleLoadFile(file: File) {
    if (!session) { setError('กรุณาเข้าสู่ระบบก่อน'); return; }
    setError(''); setStatus(`กำลังอ่าน ${file.name}`);
    try {
      const wb = /\.(xlsx|xlsm|zip)$/i.test(file.name) ? await parseWorkbookFile(file, { ...db }, profile) : await parseCsvFile(file, profile);
      persistWorkbook(wb);
      setActiveSheet(0);
      setView('dashboard');
      setStatus(`โหลดสำเร็จ: ${wb.sheets.length} sheet`);
    } catch (e) {
      setWorkbook(null);
      setError((e as Error).message || 'ไม่สามารถอ่านไฟล์ได้');
      setStatus('เกิดข้อผิดพลาด');
    }
  }

  async function handleImportDatabase(file: File) {
    if (!session) { setError('กรุณาเข้าสู่ระบบก่อน'); return; }
    try {
      const imported = parseDatabaseFile(await file.text());
      setDB(imported);
      setStatus('นำเข้าฐานข้อมูลสำเร็จ');
    } catch (e) {
      setError((e as Error).message || 'ไม่สามารถนำเข้าฐานข้อมูลได้');
    }
  }

  function handleExportDatabase() { downloadText('doit-database.json', exportDatabase(db), 'application/json;charset=utf-8'); }

  function updateMapping(field: CanonicalField, header: string) {
    if (!active || !workbook) return;
    const nextSheets = workbook.sheets.map((sheet, idx) => idx === activeSheet ? { ...sheet, mapping: { ...sheet.mapping, [field]: header || undefined } } : sheet);
    persistWorkbook(computeWorkbook(workbook.sourceName, workbook.parsedAt, nextSheets.map((s) => ({ name: s.name, headers: s.headers, rows: s.rows, mapping: s.mapping })), workbook.warnings, profile));
    setDB((prev) => ({ ...prev, templates: { ...prev.templates, [active.name]: { ...active.mapping, [field]: header || undefined } } }));
  }

  function saveTemplate() { if (!active) return; setDB((prev) => ({ ...prev, templates: { ...prev.templates, [active.name]: active.mapping } })); setStatus(`บันทึก mapping ของ ${active.name}`); }
  const styleBtn = (id: View) => ({ padding: '10px 14px', borderRadius: 999, border: `1px solid ${view === id ? t.accent : t.border}`, background: view === id ? `${t.accent}22` : t.panel, color: t.text, fontWeight: 700 as const, cursor: 'pointer' });
  const shell = { minHeight: '100vh', background: t.shell, color: t.text, fontFamily: 'Inter, system-ui, sans-serif' } as const;
  const panel = { background: t.panel, border: `1px solid ${t.border}`, borderRadius: 20, boxShadow: dark ? '0 20px 60px rgba(0,0,0,0.18)' : '0 18px 48px rgba(15,23,42,0.06)' } as const;
  const card = { background: t.card, border: `1px solid ${t.border}`, borderRadius: 16 } as const;
  function exportReport() { if (workbook) downloadText(`${workbook.sourceName.replace(/\.[^.]+$/, '')}.md`, workbook.reportMarkdown, 'text/markdown;charset=utf-8'); }
  function exportJson() { if (workbook) downloadJson(`${workbook.sourceName.replace(/\.[^.]+$/, '')}.json`, workbook); }

  return (
    <div style={shell}>
      <div style={{ maxWidth: 1480, margin: '0 auto', padding: 20 }}>
        {!session ? (
          <div style={{ ...panel, maxWidth: 720, margin: '48px auto', padding: 24 }}>
            <h1 style={{ margin: 0, fontSize: 'clamp(2rem, 4vw, 3rem)' }}>DOIT Workspace</h1>
            <p style={{ color: t.muted }}>Sign in to use dashboard, TOD/store aggregation, relations, anomaly scoring and reports.</p>
            <div style={{ display: 'grid', gap: 12, marginTop: 16 }}>
              <input id="name" placeholder="ชื่อผู้ใช้" style={{ padding: 12, borderRadius: 12, border: `1px solid ${t.border}`, background: t.card, color: t.text }} />
              <input id="org" placeholder="ทีม/องค์กร" style={{ padding: 12, borderRadius: 12, border: `1px solid ${t.border}`, background: t.card, color: t.text }} />
              <button onClick={login} style={{ padding: 12, borderRadius: 12, border: 'none', background: t.accent, color: '#082f49', fontWeight: 800, cursor: 'pointer' }}>เข้าสู่ระบบ</button>
            </div>
          </div>
        ) : (
          <>
            <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap', alignItems: 'flex-end' }}>
              <div>
                <div style={{ fontSize: 13, color: t.muted, textTransform: 'uppercase', letterSpacing: 1.2 }}>DOIT Web App</div>
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
                  <div style={{ color: t.muted }}>รองรับ .xlsx, .xlsm, .zip, .csv, database JSON</div>
                </div>
                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                  <label style={{ cursor: 'pointer', padding: '12px 16px', borderRadius: 14, background: t.accent, color: '#082f49', fontWeight: 800 }}>
                    Load file
                    <input type="file" accept=".xlsx,.xlsm,.zip,.csv,.txt" style={{ display: 'none' }} onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleLoadFile(f); }} />
                  </label>
                  <label style={{ cursor: 'pointer', padding: '12px 16px', borderRadius: 14, background: t.panel, color: t.text, border: `1px solid ${t.border}`, fontWeight: 800 }}>
                    Import DB
                    <input type="file" accept="application/json,.json" style={{ display: 'none' }} onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleImportDatabase(f); }} />
                  </label>
                </div>
              </div>
              {error && <div style={{ marginTop: 12, padding: 12, borderRadius: 12, background: `${t.danger}18`, color: t.danger }}>{error}</div>}
            </div>

            <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 16 }}>
              {(['dashboard', 'sheets', 'relations', 'tod', 'reports', 'database', 'settings'] as View[]).map((v) => <button key={v} style={styleBtn(v)} onClick={() => setView(v)}>{v}</button>)}
              <button style={styleBtn('dashboard')} onClick={exportReport} disabled={!workbook}>Export MD</button>
              <button style={styleBtn('dashboard')} onClick={exportJson} disabled={!workbook}>Export JSON</button>
              <button style={styleBtn('database')} onClick={logout}>Logout</button>
            </div>

            {workbook && view === 'dashboard' && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 16 }}>
                  {[
                    { label: 'Sheets', value: workbook.sheets.length },
                    { label: 'Rows', value: workbook.rowCount },
                    { label: 'Anomalies', value: workbook.anomalyCount },
                    { label: 'Relations', value: workbook.relations.length },
                  ].map((s) => <div key={s.label} style={{ ...card, padding: 18 }}><div style={{ color: t.muted, fontSize: 12 }}>{s.label}</div><div style={{ fontSize: 28, fontWeight: 900, marginTop: 4 }}>{s.value}</div></div>)}
                </div>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>Top anomalies</div>
                  <div style={{ marginTop: 12, display: 'grid', gap: 10 }}>
                    {allAnomalies.slice(0, 12).map((a) => <div key={`${a.sheet}-${a.rowIndex}`} style={{ padding: 12, borderRadius: 12, background: card.background, border: `1px solid ${t.border}` }}><div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}><strong>{a.sheet} · {a.invoiceNo || `Row ${a.rowIndex + 1}`}</strong><span style={{ color: t.danger, fontWeight: 800 }}>score {a.score}</span></div><div style={{ color: t.muted, marginTop: 4, fontSize: 12 }}>{a.reasons.join(' · ') || 'normal'}</div></div>)}
                  </div>
                </div>
              </div>
            )}

            {workbook && view === 'sheets' && active && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>{workbook.sheets.map((sheet, idx) => <button key={sheet.name + idx} onClick={() => setActiveSheet(idx)} style={{ ...styleBtn('sheets'), background: idx === activeSheet ? `${t.accent}22` : t.panel }}>{sheet.name}</button>)}</div>
                <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1.6fr) minmax(280px, 0.9fr)', gap: 18 }}>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ ...panel, padding: 18 }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                        <div><div style={{ fontSize: 12, color: t.muted }}>Active sheet</div><div style={{ fontSize: 20, fontWeight: 800 }}>{active.name}</div></div>
                        <div style={{ color: t.muted }}>{active.rows.length} rows · {active.headers.length} columns</div>
                      </div>
                      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 14 }}>
                        <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search..." style={{ flex: '1 1 220px', padding: 12, borderRadius: 12, border: `1px solid ${t.border}`, background: t.card, color: t.text }} />
                        <button onClick={saveTemplate} style={styleBtn('sheets')}>Save mapping</button>
                      </div>
                    </div>
                    <div style={{ marginTop: 16, overflowX: 'auto', borderRadius: 18, border: `1px solid ${t.border}`, background: t.panel }}>
                      <table style={{ width: '100%', minWidth: 900 }}>
                        <thead><tr>{active.headers.map((h) => <th key={h} style={{ position: 'sticky', top: 0, background: t.panel, textAlign: 'left', padding: 12, borderBottom: `1px solid ${t.border}`, color: t.muted }}>{h}</th>)}</tr></thead>
                        <tbody>{visibleRows.map((row, i) => <tr key={i} style={{ background: i % 2 ? t.tableAlt : 'transparent' }}>{active.headers.map((h) => <td key={h} style={{ padding: 12, borderBottom: `1px solid ${t.border}` }}>{safeText(row[h]) || '—'}</td>)}</tr>)}</tbody>
                      </table>
                    </div>
                  </div>
                  <div style={{ display: 'grid', gap: 18 }}>
                    <div style={{ ...panel, padding: 18 }}>
                      <div style={{ fontWeight: 800, fontSize: 18 }}>Auto mapping template</div>
                      <div style={{ display: 'grid', gap: 10, marginTop: 12 }}>
                        {(Object.keys(makeTemplate(active.headers)) as CanonicalField[]).map((field) => (
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
                      <div style={{ fontWeight: 800, fontSize: 18 }}>Summary</div>
                      <div style={{ display: 'grid', gap: 8, marginTop: 12 }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Rows</span><strong>{active.rows.length}</strong></div>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Anomalies</span><strong>{active.anomalies.length}</strong></div>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Stores</span><strong>{active.storeSummary.length}</strong></div>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>TOD buckets</span><strong>{active.todSummary.length}</strong></div>
                      </div>
                    </div>
                    <div style={{ ...panel, padding: 18 }}>
                      <div style={{ fontWeight: 800, fontSize: 18 }}>Top numeric columns</div>
                      <div style={{ display: 'grid', gap: 10, marginTop: 12 }}>{active.numericSummary.slice(0, 8).map((n) => <div key={n.name} style={{ padding: 12, borderRadius: 12, background: t.card, border: `1px solid ${t.border}` }}><div style={{ color: t.muted, fontSize: 12 }}>{n.name}</div><strong>{Number.isInteger(n.total) ? n.total.toLocaleString() : n.total.toFixed(2)}</strong><div style={{ color: t.muted, fontSize: 12 }}>{n.count} numeric cells</div></div>)}</div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {workbook && view === 'relations' && (
              <div style={{ marginTop: 22, ...panel, padding: 18 }}>
                <div style={{ fontWeight: 800, fontSize: 18 }}>Cross-sheet relation engine</div>
                <div style={{ marginTop: 12, display: 'grid', gap: 10 }}>{workbook.relations.length ? workbook.relations.map((r, idx) => <div key={idx} style={{ padding: 12, borderRadius: 12, background: card.background, border: `1px solid ${t.border}` }}><strong>{r.leftSheet} ↔ {r.rightSheet}</strong><div style={{ color: t.muted, fontSize: 12 }}>overlap: {r.overlap} · left: {r.leftOnly} · right: {r.rightOnly}</div></div>) : <div style={{ color: t.muted }}>No strong cross-sheet relation detected</div>}</div>
              </div>
            )}

            {workbook && view === 'tod' && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>TOD / Store aggregation</div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 16, marginTop: 12 }}>{allStores.slice(0, 8).map((s, idx) => <div key={idx} style={{ ...card, padding: 14 }}><div style={{ fontWeight: 800 }}>{s.store}</div><div style={{ color: t.muted, fontSize: 12 }}>{s.sheet}</div><div style={{ marginTop: 8 }}>count {s.count}</div><div>amount {Number.isInteger(s.amount) ? s.amount.toLocaleString() : s.amount.toFixed(2)}</div><div style={{ color: t.danger }}>suspicious {s.suspicious}</div></div>)}</div>
                </div>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>TOD buckets</div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 12, marginTop: 12 }}>{allTod.map((x, idx) => <div key={idx} style={{ ...card, padding: 14 }}><div style={{ fontWeight: 800 }}>{x.bucket}</div><div style={{ color: t.muted, fontSize: 12 }}>{x.sheet}</div><div>count {x.count}</div><div>amount {Number.isInteger(x.amount) ? x.amount.toLocaleString() : x.amount.toFixed(2)}</div></div>)}</div>
                </div>
              </div>
            )}

            {workbook && view === 'reports' && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                    <div><div style={{ fontWeight: 800, fontSize: 18 }}>Report generator</div><div style={{ color: t.muted, marginTop: 4 }}>Markdown summary of workbook, anomalies and relations</div></div>
                    <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}><button onClick={exportReport} style={styleBtn('reports')}>Download MD</button><button onClick={exportJson} style={styleBtn('reports')}>Download JSON</button></div>
                  </div>
                  <pre style={{ whiteSpace: 'pre-wrap', marginTop: 16, padding: 16, borderRadius: 14, background: t.card, border: `1px solid ${t.border}`, overflowX: 'auto' }}>{workbook.reportMarkdown}</pre>
                </div>
              </div>
            )}

            {view === 'database' && (
              <div style={{ marginTop: 22, display: 'grid', gap: 18 }}>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>Persistent local database</div>
                  <div style={{ color: t.muted, marginTop: 4 }}>Sessions, workbooks, reports and templates are saved in browser storage.</div>
                  <div style={{ display: 'grid', gap: 10, marginTop: 12 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Sessions</span><strong>{db.sessions.length}</strong></div>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Workbooks</span><strong>{db.workbooks.length}</strong></div>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Reports</span><strong>{db.reports.length}</strong></div>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}><span style={{ color: t.muted }}>Templates</span><strong>{Object.keys(db.templates).length}</strong></div>
                  </div>
                  <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 14 }}>
                    <button onClick={handleExportDatabase} style={styleBtn('database')}>Export DB</button>
                    <label style={{ cursor: 'pointer', padding: '10px 14px', borderRadius: 12, border: `1px solid ${t.border}`, background: t.panel, color: t.text, fontWeight: 700 }}>
                      Import DB
                      <input type="file" accept="application/json,.json" style={{ display: 'none' }} onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleImportDatabase(f); }} />
                    </label>
                    <button onClick={() => { setDB(DEFAULT_DB); setStatus('Database cleared'); }} style={styleBtn('database')}>Clear DB</button>
                  </div>
                </div>
                <div style={{ ...panel, padding: 18 }}>
                  <div style={{ fontWeight: 800, fontSize: 18 }}>Session</div>
                  <div style={{ marginTop: 10 }}>User: <strong>{session.name}</strong></div>
                  <div>Org: <strong>{session.org}</strong></div>
                  <div>Token: <code>{session.token.slice(0, 8)}...</code></div>
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
