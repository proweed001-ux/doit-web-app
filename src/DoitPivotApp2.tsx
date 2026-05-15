import JSZip from 'jszip';
import { useMemo, useState } from 'react';

type Cell = string | number | boolean | null;
type Row = Record<string, Cell>;
type Sheet = { name: string; headers: string[]; rows: Row[] };
type MapKey = 'ps' | 'brand' | 'size' | 'sku' | 'desc' | 'type' | 'qty' | 'amt' | 'date';
type FieldMap = Partial<Record<MapKey, string>>;
type Line = { rowNo: number; ps: string; brand: string; size: string; sku: string; desc: string; type: string; date: string; qty: number; amt: number };
type Pivot = Omit<Line, 'rowNo' | 'date'> & { lineCount: number };

const aliases: Record<MapKey, string[]> = {
  ps: ['sosellergroup', 'sellergroup', 'ps', 'psname', 'salesperson', 'salesgroup'],
  brand: ['brand', 'brandname'],
  size: ['tassizegroup', 'tasizegroup', 'sizegroup', 'tasgroup'],
  sku: ['sku', 'skucode', 'itemcode', 'productcode', 'material', 'materialcode'],
  desc: ['description', 'desc', 'skudescription', 'productdescription', 'itemdescription'],
  type: ['sotypeid', 'sotype', 'sotypecode', 'doctype', 'documenttype'],
  qty: ['shipqtypcs', 'shipqtypc', 'shipqty', 'qtypcs'],
  amt: ['invoiceamt', 'invoiceamount'],
  date: ['sodate', 'invoicedate', 'docdate', 'documentdate', 'date'],
};

const labels: Record<MapKey, string> = {
  ps: 'SO Seller Group / PS',
  brand: 'Brand',
  size: 'TAS_SizeGroup',
  sku: 'SKU',
  desc: 'Description',
  type: 'SOTypeID',
  qty: 'ShipQtyPCS',
  amt: 'InvoiceAmt',
  date: 'SO_Date / Invoice_Date',
};

const required: MapKey[] = ['ps', 'brand', 'size', 'sku', 'desc', 'type', 'qty', 'amt'];
const maxFileBytes = 35 * 1024 * 1024;

function text(v: unknown): string {
  return String(v ?? '').replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim();
}

function norm(v: unknown): string {
  return text(v).toLowerCase().normalize('NFKD').replace(/[^\p{L}\p{N}]+/gu, '');
}

function num(v: Cell): number {
  if (typeof v === 'number') return Number.isFinite(v) ? v : Number.NaN;
  const raw = text(v);
  if (!raw) return Number.NaN;
  const paren = /^\(.*\)$/.test(raw);
  const n = Number(raw.replace(/[(),\s]/g, '').replace(/−/g, '-'));
  return Number.isFinite(n) ? (paren && n > 0 ? -n : n) : Number.NaN;
}

function fmt(n: number, digits = 2): string {
  return Number.isFinite(n) ? n.toLocaleString('th-TH', { maximumFractionDigits: digits }) : '0';
}

function colIndex(ref: string): number {
  const m = /([A-Z]+)/i.exec(ref);
  return m ? m[1].toUpperCase().split('').reduce((s, c) => s * 26 + c.charCodeAt(0) - 64, 0) - 1 : 0;
}

function colName(i: number): string {
  let n = i + 1, out = '';
  while (n > 0) { const m = (n - 1) % 26; out = String.fromCharCode(65 + m) + out; n = Math.floor((n - m) / 26); }
  return out;
}

function splitCsv(line: string): string[] {
  const out: string[] = [];
  let cur = '', q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') { if (q && line[i + 1] === '"') { cur += '"'; i++; } else q = !q; continue; }
    if (ch === ',' && !q) { out.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur.trim());
  return out;
}

function uniqueHeaders(cells: Cell[]): string[] {
  const seen = new Map<string, number>();
  return cells.map((v, i) => {
    const base = text(v) || `Column ${colName(i)}`;
    const c = seen.get(base) ?? 0;
    seen.set(base, c + 1);
    return c ? `${base} (${c + 1})` : base;
  });
}

function detect(headers: string[]): FieldMap {
  const byNorm = new Map(headers.map((h) => [norm(h), h]));
  const out: FieldMap = {};
  (Object.keys(aliases) as MapKey[]).forEach((k) => {
    const hit = aliases[k].map((a) => byNorm.get(a)).find(Boolean);
    if (hit) out[k] = hit;
  });
  return out;
}

function scoreHeader(row: Cell[]): number {
  const map = detect(row.map(text));
  return required.reduce((s, k) => s + (map[k] ? 1 : 0), 0);
}

function headerRow(matrix: Cell[][]): number {
  let best = Math.max(0, matrix.findIndex((r) => r?.some((c) => text(c))));
  let score = -1;
  matrix.forEach((r, i) => { const s = scoreHeader(r ?? []); if (s > score) { best = i; score = s; } });
  return best;
}

function sharedStrings(xml: string): string[] {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return Array.from(doc.getElementsByTagName('si')).map((el) => text(el.textContent));
}

function workbookRefs(wbXml: string, relsXml: string): { name: string; path: string }[] {
  const wb = new DOMParser().parseFromString(wbXml, 'application/xml');
  const rels = new DOMParser().parseFromString(relsXml, 'application/xml');
  const relMap = new Map<string, string>();
  Array.from(rels.getElementsByTagName('Relationship')).forEach((r) => {
    const id = r.getAttribute('Id'), target = r.getAttribute('Target');
    if (id && target) relMap.set(id, target.startsWith('/') ? target.slice(1) : target);
  });
  return Array.from(wb.getElementsByTagName('sheet')).map((s, i) => {
    const name = s.getAttribute('name') ?? `Sheet${i + 1}`;
    const rid = s.getAttribute('r:id') ?? '';
    const target = relMap.get(rid) ?? `worksheets/sheet${i + 1}.xml`;
    return { name, path: target.startsWith('xl/') ? target : `xl/${target}` };
  });
}

function sheetFromXml(xml: string, ss: string[], name: string): Sheet {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const matrix: Cell[][] = [];
  Array.from(doc.getElementsByTagName('row')).forEach((row) => {
    const ri = Number(row.getAttribute('r') ?? matrix.length + 1) - 1;
    matrix[ri] ??= [];
    Array.from(row.getElementsByTagName('c')).forEach((cell) => {
      const ci = cell.getAttribute('r') ? colIndex(cell.getAttribute('r') ?? '') : matrix[ri].length;
      const type = cell.getAttribute('t') ?? '';
      const v = cell.getElementsByTagName('v')[0];
      const inline = cell.getElementsByTagName('is')[0];
      let value: Cell = '';
      if (type === 's') value = ss[Number(text(v?.textContent))] ?? '';
      else if (type === 'inlineStr') value = text(inline?.textContent);
      else if (type === 'b') value = text(v?.textContent) === '1';
      else { const raw = text(v?.textContent ?? cell.textContent); value = raw === '' ? '' : Number.isFinite(Number(raw)) ? Number(raw) : raw; }
      matrix[ri][ci] = value;
    });
  });
  const hi = headerRow(matrix);
  const headers = uniqueHeaders(matrix[hi] ?? []);
  const rows = matrix.slice(hi + 1).filter((r) => r?.some((c) => text(c))).map((r) => {
    const obj: Row = {};
    headers.forEach((h, i) => { obj[h] = r?.[i] ?? ''; });
    return obj;
  });
  return { name, headers, rows };
}

function sheetFromCsv(raw: string, name: string): Sheet {
  const matrix = raw.split(/\r?\n/).filter((l) => text(l)).map(splitCsv);
  const hi = headerRow(matrix);
  const headers = uniqueHeaders(matrix[hi] ?? []);
  const rows = matrix.slice(hi + 1).map((cells) => {
    const obj: Row = {};
    headers.forEach((h, i) => { obj[h] = cells[i] ?? ''; });
    return obj;
  });
  return { name, headers, rows };
}

async function parseFile(file: File): Promise<{ name: string; sheets: Sheet[]; warnings: string[] }> {
  if (/\.(csv|txt)$/i.test(file.name)) return { name: file.name, sheets: [sheetFromCsv(await file.text(), file.name)], warnings: ['อ่านเป็น CSV/TXT'] };
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const wb = await zip.file('xl/workbook.xml')?.async('text');
  const rels = await zip.file('xl/_rels/workbook.xml.rels')?.async('text');
  if (!wb || !rels) throw new Error('ไฟล์นี้ไม่ใช่ Excel มาตรฐาน หรือไฟล์เสีย');
  const ss = sharedStrings(await zip.file('xl/sharedStrings.xml')?.async('text') ?? '');
  const sheets: Sheet[] = [];
  const warnings: string[] = [];
  for (const ref of workbookRefs(wb, rels)) {
    const xml = await zip.file(ref.path)?.async('text');
    if (!xml) { warnings.push(`อ่าน sheet ${ref.name} ไม่ได้`); continue; }
    sheets.push(sheetFromXml(xml, ss, ref.name));
  }
  return { name: file.name, sheets, warnings };
}

function val(row: Row, map: FieldMap, key: MapKey): Cell {
  const h = map[key];
  return h ? row[h] ?? '' : '';
}

function rfc(type: string): boolean { return norm(type) === 'rfc' || /\bRFC\b/i.test(type); }
function signed(n: number, type: string): number { return rfc(type) && n > 0 ? -n : Number.isFinite(n) ? n : 0; }
function excelDate(v: number): string { return v > 20000 && v < 70000 ? new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000).toISOString().slice(0, 10) : ''; }
function dateVal(v: Cell): string { return typeof v === 'number' ? excelDate(v) || text(v) : text(v); }

function linesFrom(sheet: Sheet, map: FieldMap): Line[] {
  if (!map.qty || !map.amt) return [];
  return sheet.rows.flatMap((row, i) => {
    const type = text(val(row, map, 'type'));
    const sku = text(val(row, map, 'sku'));
    const desc = text(val(row, map, 'desc'));
    const qty = num(val(row, map, 'qty'));
    const amt = num(val(row, map, 'amt'));
    const hasIdentity = Boolean(sku || desc || text(val(row, map, 'ps')) || text(val(row, map, 'brand')));
    const hasValue = Number.isFinite(qty) || Number.isFinite(amt);
    if (!hasIdentity || !hasValue) return [];
    return [{
      rowNo: i + 2,
      ps: text(val(row, map, 'ps')) || 'ไม่ระบุ PS',
      brand: text(val(row, map, 'brand')) || 'ไม่ระบุ Brand',
      size: text(val(row, map, 'size')) || 'ไม่ระบุ Size',
      sku: sku || 'ไม่ระบุ SKU',
      desc: desc || 'ไม่ระบุ Description',
      type: type || 'ไม่ระบุ SOTypeID',
      date: dateVal(val(row, map, 'date')) || 'ไม่ระบุวันที่',
      qty: signed(qty, type),
      amt: signed(amt, type),
    }];
  });
}

function pivot(lines: Line[]): Pivot[] {
  const map = new Map<string, Pivot>();
  for (const line of lines) {
    const key = JSON.stringify([line.ps, line.brand, line.size, line.sku, line.desc, line.type]);
    const cur = map.get(key) ?? { ps: line.ps, brand: line.brand, size: line.size, sku: line.sku, desc: line.desc, type: line.type, qty: 0, amt: 0, lineCount: 0 };
    cur.qty += line.qty;
    cur.amt += line.amt;
    cur.lineCount += 1;
    map.set(key, cur);
  }
  return Array.from(map.values()).sort((a, b) => a.ps.localeCompare(b.ps) || a.brand.localeCompare(b.brand) || a.size.localeCompare(b.size) || a.sku.localeCompare(b.sku) || a.type.localeCompare(b.type));
}

function sumQty(rows: { qty: number }[]) { return rows.reduce((s, r) => s + r.qty, 0); }
function sumAmt(rows: { amt: number }[]) { return rows.reduce((s, r) => s + r.amt, 0); }
function unique(values: string[]) { return Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b)); }
function csv(v: unknown) { const s = text(v); return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; }
function download(name: string, body: string, type = 'text/plain;charset=utf-8') { const a = document.createElement('a'); const url = URL.createObjectURL(new Blob([body], { type })); a.href = url; a.download = name; a.click(); URL.revokeObjectURL(url); }

function warnings(sheet: Sheet | null, map: FieldMap, parseWarnings: string[]): string[] {
  if (!sheet) return parseWarnings;
  const out = [...parseWarnings];
  required.forEach((k) => { if (!map[k]) out.push(`ไม่พบคอลัมน์จำเป็น: ${labels[k]}`); });
  const bad = sheet.headers.filter((h) => /tot.*invc|invoice.*total|total.*invoice|header.*amount/i.test(h) || ['totinvc', 'totalinvoice', 'invoicetotal'].includes(norm(h)));
  if (bad.length) out.push(`พบคอลัมน์ยอดหัวบิลที่ไม่ใช้: ${bad.join(', ')} — app ใช้เฉพาะ InvoiceAmt ต่อ line`);
  return out;
}

export default function DoitPivotApp2() {
  const [book, setBook] = useState<{ name: string; sheets: Sheet[]; warnings: string[] } | null>(null);
  const [sheetName, setSheetName] = useState('');
  const [ps, setPs] = useState('ALL');
  const [date, setDate] = useState('ALL');
  const [type, setType] = useState('ALL');
  const [q, setQ] = useState('');
  const [status, setStatus] = useState('พร้อมอ่านไฟล์ DOIT');
  const [error, setError] = useState('');

  const sheet = useMemo(() => book?.sheets.find((s) => s.name === sheetName) ?? book?.sheets[0] ?? null, [book, sheetName]);
  const map = useMemo(() => sheet ? detect(sheet.headers) : {}, [sheet]);
  const rawLines = useMemo(() => sheet ? linesFrom(sheet, map) : [], [sheet, map]);
  const filtered = useMemo(() => rawLines.filter((l) => {
    if (ps !== 'ALL' && l.ps !== ps) return false;
    if (date !== 'ALL' && l.date !== date) return false;
    if (type !== 'ALL' && l.type !== type) return false;
    const nq = norm(q);
    return !nq || [l.ps, l.brand, l.size, l.sku, l.desc, l.type, l.date].some((x) => norm(x).includes(nq));
  }), [rawLines, ps, date, type, q]);
  const piv = useMemo(() => pivot(filtered), [filtered]);
  const warn = useMemo(() => warnings(sheet, map, book?.warnings ?? []), [sheet, map, book]);
  const ready = required.every((k) => Boolean(map[k]));

  async function load(file: File) {
    if (file.size > maxFileBytes) { setError('ไฟล์ใหญ่เกิน 35 MB'); return; }
    setError(''); setStatus(`กำลังอ่าน ${file.name}`);
    try {
      const parsed = await parseFile(file);
      const first = parsed.sheets.find((s) => { const m = detect(s.headers); return m.qty && m.amt; }) ?? parsed.sheets[0];
      setBook(parsed); setSheetName(first?.name ?? ''); setPs('ALL'); setDate('ALL'); setType('ALL'); setQ(''); setStatus(`อ่านสำเร็จ ${parsed.sheets.length} sheet`);
    } catch (e) { setBook(null); setError((e as Error).message || 'อ่านไฟล์ไม่สำเร็จ'); setStatus('อ่านไฟล์ไม่สำเร็จ'); }
  }

  function exportCsv() {
    const head = ['SO Seller Group', 'Brand', 'TAS_SizeGroup', 'SKU', 'Description', 'SOTypeID', 'ShipQtyPCS', 'InvoiceAmt', 'LineCount'];
    const body = piv.map((r) => [r.ps, r.brand, r.size, r.sku, r.desc, r.type, r.qty, r.amt, r.lineCount].map(csv).join(','));
    download('doit-pivot-result.csv', `\uFEFF${[head.join(','), ...body].join('\n')}`, 'text/csv;charset=utf-8');
  }

  const card = { background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: 16, padding: 16 };
  const panel = { background: '#fff', border: '1px solid #cbd5e1', borderRadius: 20, padding: 20, boxShadow: '0 18px 48px rgba(15,23,42,.06)' };
  const btn = { border: 0, borderRadius: 14, padding: '16px 20px', background: '#0369a1', color: '#fff', fontWeight: 900, fontSize: 18, cursor: 'pointer' };
  const input = { padding: 14, borderRadius: 12, border: '1px solid #cbd5e1', fontSize: 17, background: '#fff', color: '#0f172a' };

  return <div style={{ minHeight: '100vh', background: '#f8fafc', color: '#0f172a', fontFamily: 'Tahoma, Noto Sans Thai, system-ui, sans-serif', fontSize: 18 }}><div style={{ maxWidth: 1500, margin: '0 auto', padding: 22 }}>
    <div style={{ ...panel, display: 'flex', justifyContent: 'space-between', gap: 18, flexWrap: 'wrap', alignItems: 'center' }}>
      <div><div style={{ color: '#475569', fontWeight: 900 }}>DOIT Pivot App ใหม่</div><h1 style={{ margin: '8px 0', fontSize: 'clamp(2.1rem,4vw,3.6rem)' }}>ตรวจยอดตาม Pivot จริง</h1><div style={{ color: '#475569' }}>source grain = invoice line · filter ก่อน group · SUM ShipQtyPCS และ InvoiceAmt เท่านั้น</div></div>
      <label style={btn}>เลือกไฟล์ DOIT<input type="file" accept=".xlsx,.xlsm,.csv,.txt" style={{ display: 'none' }} onChange={(e) => { const f = e.target.files?.[0]; if (f) void load(f); }} /></label>
    </div>
    <div style={{ ...panel, marginTop: 16, padding: 16 }}><b>สถานะ:</b> {status}{book && <span style={{ color: '#475569' }}> · {book.name}</span>}{error && <div style={{ color: '#b91c1c', fontWeight: 900, marginTop: 8 }}>{error}</div>}</div>

    {!book && <div style={{ marginTop: 16, display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(240px,1fr))', gap: 14 }}>{[['1','เลือกไฟล์ DOIT'],['2','อ่าน raw invoice line'],['3','filter PS/วัน/ประเภท'],['4','ดู Pivot/โหลด CSV']].map(([n, s]) => <div key={n} style={card}><div style={{ fontSize: 34, fontWeight: 900, color: '#0369a1' }}>{n}</div><b>{s}</b></div>)}</div>}

    {book && sheet && <>
      <div style={{ marginTop: 16, display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(190px,1fr))', gap: 14 }}>{[['Raw lines หลัง filter', filtered.length], ['Pivot rows', piv.length], ['SUM ShipQtyPCS', sumQty(filtered)], ['SUM InvoiceAmt', sumAmt(filtered)], ['RFC amount', sumAmt(filtered.filter((l) => rfc(l.type)))]].map(([k, v]) => <div key={String(k)} style={card}><div style={{ color: '#475569' }}>{k}</div><div style={{ fontSize: 30, fontWeight: 900 }}>{typeof v === 'number' ? fmt(v) : v}</div></div>)}</div>
      <div style={{ ...panel, marginTop: 16 }}><div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(180px,1fr))', gap: 12 }}>
        <label>Sheet<select value={sheet.name} onChange={(e) => setSheetName(e.target.value)} style={{ ...input, width: '100%' }}>{book.sheets.map((s) => <option key={s.name}>{s.name}</option>)}</select></label>
        <label>PS<select value={ps} onChange={(e) => setPs(e.target.value)} style={{ ...input, width: '100%' }}><option value="ALL">ทุก PS</option>{unique(rawLines.map((l) => l.ps)).map((x) => <option key={x}>{x}</option>)}</select></label>
        <label>วันที่<select value={date} onChange={(e) => setDate(e.target.value)} style={{ ...input, width: '100%' }}><option value="ALL">ทุกวันที่</option>{unique(rawLines.map((l) => l.date)).map((x) => <option key={x}>{x}</option>)}</select></label>
        <label>SOTypeID<select value={type} onChange={(e) => setType(e.target.value)} style={{ ...input, width: '100%' }}><option value="ALL">ทุกประเภท</option>{unique(rawLines.map((l) => l.type)).map((x) => <option key={x}>{x}</option>)}</select></label>
        <label>ค้นหา<input value={q} onChange={(e) => setQ(e.target.value)} style={{ ...input, width: '100%', boxSizing: 'border-box' }} /></label>
      </div><div style={{ marginTop: 14, display: 'flex', gap: 10, flexWrap: 'wrap' }}><button style={btn} onClick={exportCsv} disabled={!piv.length}>โหลด Pivot CSV</button><button style={{ ...btn, background: '#fff', color: '#0f172a', border: '1px solid #cbd5e1' }} onClick={() => setBook(null)}>เริ่มใหม่</button></div></div>
      <div style={{ ...panel, marginTop: 16 }}><h2 style={{ marginTop: 0 }}>Field mapping</h2><div style={{ color: ready ? '#047857' : '#b91c1c', fontWeight: 900 }}>{ready ? 'พร้อมคำนวณแบบ Pivot' : 'ขาด field สำคัญ'}</div><div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(230px,1fr))', gap: 10, marginTop: 10 }}>{required.map((k) => <div key={k} style={card}><div style={{ color: '#475569' }}>{labels[k]}</div><b>{map[k] ?? 'ไม่พบ'}</b></div>)}</div>{warn.map((w, i) => <div key={i} style={{ marginTop: 8, padding: 10, borderRadius: 10, background: '#fef2f2', color: '#b91c1c', fontWeight: 800 }}>{w}</div>)}</div>
      <div style={{ ...panel, marginTop: 16 }}><h2 style={{ marginTop: 0 }}>Pivot result</h2><div style={{ overflowX: 'auto' }}><table style={{ width: '100%', minWidth: 1100, borderCollapse: 'collapse' }}><thead><tr>{['SO Seller Group','Brand','TAS_SizeGroup','SKU','Description','SOTypeID','ShipQtyPCS','InvoiceAmt','LineCount'].map((h) => <th key={h} style={{ textAlign: 'left', padding: 12, borderBottom: '1px solid #cbd5e1', background: '#f1f5f9' }}>{h}</th>)}</tr></thead><tbody>{piv.slice(0,250).map((r, i) => <tr key={i}>{[r.ps,r.brand,r.size,r.sku,r.desc,r.type,fmt(r.qty),fmt(r.amt),fmt(r.lineCount,0)].map((x, j) => <td key={j} style={{ padding: 12, borderBottom: '1px solid #cbd5e1', textAlign: j >= 6 ? 'right' : 'left' }}>{x}</td>)}</tr>)}</tbody></table></div></div>
      <div style={{ ...panel, marginTop: 16 }}><h2 style={{ marginTop: 0 }}>Raw invoice lines หลัง filter</h2><div style={{ overflowX: 'auto' }}><table style={{ width: '100%', minWidth: 1050, borderCollapse: 'collapse' }}><thead><tr>{['Row','Date','PS','Brand','Size','SKU','Description','SOTypeID','ShipQtyPCS','InvoiceAmt'].map((h) => <th key={h} style={{ textAlign: 'left', padding: 12, borderBottom: '1px solid #cbd5e1', background: '#f1f5f9' }}>{h}</th>)}</tr></thead><tbody>{filtered.slice(0,80).map((l, i) => <tr key={i}>{[l.rowNo,l.date,l.ps,l.brand,l.size,l.sku,l.desc,l.type,fmt(l.qty),fmt(l.amt)].map((x, j) => <td key={j} style={{ padding: 12, borderBottom: '1px solid #cbd5e1', textAlign: j >= 8 ? 'right' : 'left' }}>{x}</td>)}</tr>)}</tbody></table></div></div>
    </>}
  </div></div>;
}
