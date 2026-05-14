import JSZip from 'jszip';
import { useMemo, useState } from 'react';

type CellValue = string | number | boolean | null;
type RowObject = Record<string, CellValue>;

type SheetData = {
  name: string;
  headers: string[];
  rows: RowObject[];
  rawRowCount: number;
};

type WorkbookData = {
  sheets: SheetData[];
  sourceName: string;
  warningCount: number;
  inflationHits: number;
};

function cleanText(value: unknown): string {
  return String(value ?? '')
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
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
  if (!node) return '';
  return cleanText(node.textContent ?? '');
}

function parseSharedStrings(xmlTextValue: string): string[] {
  if (!xmlTextValue) return [];
  const doc = new DOMParser().parseFromString(xmlTextValue, 'application/xml');
  return Array.from(doc.getElementsByTagName('si')).map((si) => cleanText(si.textContent ?? ''));
}

function parseWorkbookSheets(workbookXml: string, relsXml: string): { name: string; path: string }[] {
  const workbookDoc = new DOMParser().parseFromString(workbookXml, 'application/xml');
  const relsDoc = new DOMParser().parseFromString(relsXml, 'application/xml');
  const relMap = new Map<string, string>();

  Array.from(relsDoc.getElementsByTagName('Relationship')).forEach((rel) => {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (id && target) {
      relMap.set(id, target.startsWith('/') ? target.slice(1) : target);
    }
  });

  return Array.from(workbookDoc.getElementsByTagName('sheet')).map((sheet) => {
    const name = sheet.getAttribute('name') ?? 'Sheet';
    const rid = sheet.getAttribute('r:id') ?? sheet.getAttribute('id') ?? '';
    const target = relMap.get(rid) ?? 'worksheets/sheet1.xml';
    const normalized = target.startsWith('xl/') ? target : `xl/${target}`;
    return { name, path: normalized };
  });
}

function parseWorksheet(xmlTextValue: string, sharedStrings: string[], sheetName: string): SheetData {
  const doc = new DOMParser().parseFromString(xmlTextValue, 'application/xml');
  const rows = Array.from(doc.getElementsByTagName('row'));
  const matrix: Array<Array<CellValue>> = [];

  for (const row of rows) {
    const rowIndex = Number(row.getAttribute('r') ?? matrix.length + 1) - 1;
    if (!matrix[rowIndex]) matrix[rowIndex] = [];

    const cells = Array.from(row.getElementsByTagName('c'));
    for (const cell of cells) {
      const ref = cell.getAttribute('r') ?? '';
      const colIndex = ref ? colToIndex(ref) : matrix[rowIndex].length;
      const type = cell.getAttribute('t') ?? '';
      const v = cell.getElementsByTagName('v')[0];
      const inline = cell.getElementsByTagName('is')[0];
      let value: CellValue = null;

      if (type === 's') {
        const idx = Number(xmlText(v));
        value = sharedStrings[idx] ?? '';
      } else if (type === 'inlineStr') {
        value = cleanText(inline?.textContent ?? '');
      } else if (type === 'b') {
        value = xmlText(v) === '1';
      } else {
        const raw = xmlText(v);
        if (raw === '') value = '';
        else if (!Number.isNaN(Number(raw)) && raw.trim() !== '') value = Number(raw);
        else value = raw;
      }

      matrix[rowIndex][colIndex] = value;
    }
  }

  const firstNonEmptyRow = matrix.findIndex((row) => row?.some((v) => cleanText(v).length > 0));
  const headerRow = firstNonEmptyRow >= 0 ? matrix[firstNonEmptyRow] : [];
  const headers = (headerRow ?? []).map((v, i) => cleanText(v) || `Column ${indexToCol(i)}`);

  const bodyRows = matrix.slice(firstNonEmptyRow + 1).filter((row) => row?.some((v) => cleanText(v).length > 0));
  const records: RowObject[] = bodyRows.map((row) => {
    const obj: RowObject = {};
    headers.forEach((header, i) => {
      const cell = row?.[i];
      obj[header] = cell ?? '';
    });
    return obj;
  });

  return {
    name: sheetName,
    headers,
    rows: records,
    rawRowCount: records.length,
  };
}

function sumNumericColumns(rows: RowObject[]): Array<{ name: string; total: number; count: number }> {
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
    .map(([name, v]) => ({ name, total: v.total, count: v.count }))
    .sort((a, b) => Math.abs(b.total) - Math.abs(a.total))
    .slice(0, 8);
}

function detectInflation(rows: RowObject[], headers: string[]): number {
  const invoiceKey = headers.find((h) => /invoice\s*no|invoice_no|invoiceno|inv\.?\s*no|bill\s*no/i.test(h))
    ?? headers.find((h) => /invoice/i.test(h));
  const amountKey = headers.find((h) => /invoice\s*amt|invoiceamt|amount|total|ยอด|เงิน|ราคา|value/i.test(h));
  const skuKey = headers.find((h) => /sku|item|product|code/i.test(h));
  if (!invoiceKey || !amountKey) return 0;

  const byInvoice = new Map<string, { amounts: Set<number>; skus: Set<string>; count: number }>();
  for (const row of rows) {
    const inv = cleanText(row[invoiceKey]);
    if (!inv) continue;
    const rawAmount = row[amountKey];
    const amount = typeof rawAmount === 'number' ? rawAmount : Number(String(rawAmount).replace(/,/g, ''));
    if (!Number.isFinite(amount)) continue;
    const sku = skuKey ? cleanText(row[skuKey]) : '';
    const current = byInvoice.get(inv) ?? { amounts: new Set<number>(), skus: new Set<string>(), count: 0 };
    current.amounts.add(amount);
    if (sku) current.skus.add(sku);
    current.count += 1;
    byInvoice.set(inv, current);
  }

  let hits = 0;
  for (const v of byInvoice.values()) {
    if (v.count >= 2 && v.amounts.size === 1 && v.skus.size >= 2) hits += 1;
  }
  return hits;
}

async function parseWorkbookFile(file: File): Promise<WorkbookData> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const workbookXml = await zip.file('xl/workbook.xml')?.async('text');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('text');
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('text');

  if (!workbookXml || !relsXml) {
    throw new Error('ไฟล์ Excel นี้ไม่ใช่ workbook มาตรฐาน');
  }

  const sharedStrings = parseSharedStrings(sharedStringsXml ?? '');
  const sheetDefs = parseWorkbookSheets(workbookXml, relsXml);
  const sheets: SheetData[] = [];
  let warningCount = 0;

  for (const sheet of sheetDefs) {
    const xml = await zip.file(sheet.path)?.async('text');
    if (!xml) {
      warningCount += 1;
      continue;
    }
    const parsed = parseWorksheet(xml, sharedStrings, sheet.name);
    sheets.push(parsed);
  }

  const flatRows = sheets.flatMap((sheet) => sheet.rows);
  const inflationHits = sheets.reduce((acc, sheet) => acc + detectInflation(sheet.rows, sheet.headers), 0);

  return {
    sheets,
    sourceName: file.name,
    warningCount,
    inflationHits,
  };
}

export default function DoitApp() {
  const [status, setStatus] = useState('พร้อมใช้งาน');
  const [workbook, setWorkbook] = useState<WorkbookData | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);
  const [error, setError] = useState('');
  const [dragOver, setDragOver] = useState(false);

  const active = workbook?.sheets[activeSheet] ?? null;
  const visibleRows = active?.rows.slice(0, 50) ?? [];
  const numericSummary = useMemo(() => active ? sumNumericColumns(active.rows) : [], [active]);

  async function loadFile(file: File) {
    setError('');
    setStatus(`กำลังอ่าน ${file.name}`);
    try {
      if (/\.(xlsx|xlsm|zip)$/i.test(file.name)) {
        const data = await parseWorkbookFile(file);
        setWorkbook(data);
        setActiveSheet(0);
        setStatus(`โหลดสำเร็จ: ${data.sheets.length} sheet`);
      } else {
        const text = await file.text();
        const lines = text.split(/\r?\n/).filter(Boolean);
        const headers = lines[0]?.split(',').map((h) => cleanText(h)) ?? [];
        const rows = lines.slice(1, 51).map((line) => {
          const cols = line.split(',');
          const obj: RowObject = {};
          headers.forEach((h, i) => { obj[h || `Column ${i + 1}`] = cleanText(cols[i] ?? ''); });
          return obj;
        });
        setWorkbook({
          sheets: [{ name: file.name, headers, rows, rawRowCount: rows.length }],
          sourceName: file.name,
          warningCount: 0,
          inflationHits: 0,
        });
        setActiveSheet(0);
        setStatus(`โหลดข้อความสำเร็จ: ${rows.length} rows`);
      }
    } catch (e) {
      setError((e as Error).message || 'ไม่สามารถอ่านไฟล์ได้');
      setStatus('เกิดข้อผิดพลาด');
      setWorkbook(null);
    }
  }

  const currentHeaders = active?.headers ?? [];

  return (
    <div style={{ minHeight: '100vh', background: '#0f172a', color: '#e2e8f0', fontFamily: 'Inter, system-ui, sans-serif' }}>
      <div style={{ maxWidth: 1280, margin: '0 auto', padding: 24 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 16, alignItems: 'flex-end', flexWrap: 'wrap' }}>
          <div>
            <div style={{ fontSize: 14, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8' }}>DOIT Web App</div>
            <h1 style={{ margin: '8px 0 4px', fontSize: 36, lineHeight: 1.1 }}>Excel Workbook Analyzer</h1>
            <div style={{ color: '#cbd5e1' }}>อัปโหลดไฟล์ .xlsx แล้วดูข้อมูลแยกชีต, สรุปตัวเลข, และตรวจความเสี่ยง invoice inflation</div>
          </div>
          <div style={{ padding: '10px 14px', borderRadius: 14, background: '#111827', border: '1px solid #334155', minWidth: 240 }}>
            <div style={{ fontSize: 12, color: '#94a3b8' }}>สถานะ</div>
            <div style={{ fontWeight: 700 }}>{status}</div>
            {workbook && <div style={{ marginTop: 6, fontSize: 12, color: '#94a3b8' }}>{workbook.sourceName}</div>}
          </div>
        </div>

        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragOver(false);
            const file = e.dataTransfer.files?.[0];
            if (file) void loadFile(file);
          }}
          style={{
            marginTop: 24,
            padding: 24,
            border: `1px dashed ${dragOver ? '#38bdf8' : '#475569'}`,
            background: dragOver ? 'rgba(14,165,233,0.08)' : '#111827',
            borderRadius: 20,
          }}
        >
          <div style={{ display: 'flex', justifyContent: 'space-between', gap: 16, alignItems: 'center', flexWrap: 'wrap' }}>
            <div>
              <div style={{ fontSize: 18, fontWeight: 700 }}>นำเข้าไฟล์ Excel</div>
              <div style={{ marginTop: 4, color: '#94a3b8' }}>รองรับไฟล์ .xlsx / .xlsm / .zip ที่มี workbook ภายใน</div>
            </div>
            <label style={{ cursor: 'pointer', padding: '10px 16px', borderRadius: 12, background: '#38bdf8', color: '#082f49', fontWeight: 700 }}>
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
          {error && <div style={{ marginTop: 16, color: '#fda4af', background: '#450a0a', padding: 12, borderRadius: 12 }}>{error}</div>}
        </div>

        {workbook && (
          <>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 16, marginTop: 24 }}>
              {[
                { label: 'Sheets', value: workbook.sheets.length },
                { label: 'Warnings', value: workbook.warningCount },
                { label: 'Inflation hits', value: workbook.inflationHits },
                { label: 'Rows', value: workbook.sheets.reduce((a, s) => a + s.rows.length, 0) },
              ].map((card) => (
                <div key={card.label} style={{ background: '#111827', border: '1px solid #334155', borderRadius: 18, padding: 18 }}>
                  <div style={{ color: '#94a3b8', fontSize: 12 }}>{card.label}</div>
                  <div style={{ fontSize: 28, fontWeight: 800, marginTop: 6 }}>{card.value}</div>
                </div>
              ))}
            </div>

            <div style={{ display: 'flex', gap: 10, marginTop: 20, flexWrap: 'wrap' }}>
              {workbook.sheets.map((sheet, idx) => (
                <button
                  key={sheet.name + idx}
                  onClick={() => setActiveSheet(idx)}
                  style={{
                    padding: '10px 14px',
                    borderRadius: 999,
                    border: idx === activeSheet ? '1px solid #38bdf8' : '1px solid #334155',
                    background: idx === activeSheet ? 'rgba(56,189,248,0.15)' : '#111827',
                    color: '#e2e8f0',
                    cursor: 'pointer',
                  }}
                >
                  {sheet.name}
                </button>
              ))}
            </div>

            {active && (
              <div style={{ marginTop: 24, display: 'grid', gap: 16 }}>
                <div style={{ background: '#111827', border: '1px solid #334155', borderRadius: 18, padding: 18 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                    <div>
                      <div style={{ color: '#94a3b8', fontSize: 12 }}>Active sheet</div>
                      <div style={{ fontSize: 20, fontWeight: 700 }}>{active.name}</div>
                    </div>
                    <div style={{ color: '#94a3b8' }}>{active.rows.length} rows · {active.headers.length} columns</div>
                  </div>
                  {numericSummary.length > 0 && (
                    <div style={{ marginTop: 16, display: 'grid', gap: 10 }}>
                      <div style={{ fontWeight: 700 }}>Top numeric columns</div>
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 10 }}>
                        {numericSummary.map((n) => (
                          <div key={n.name} style={{ border: '1px solid #334155', borderRadius: 14, padding: 12, background: '#0b1220' }}>
                            <div style={{ fontSize: 12, color: '#94a3b8' }}>{n.name}</div>
                            <div style={{ fontWeight: 800 }}>{Number.isInteger(n.total) ? n.total.toLocaleString() : n.total.toFixed(2)}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                <div style={{ overflowX: 'auto', border: '1px solid #334155', borderRadius: 18, background: '#0b1220' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 900 }}>
                    <thead>
                      <tr>
                        {currentHeaders.map((h) => (
                          <th key={h} style={{ position: 'sticky', top: 0, background: '#111827', color: '#cbd5e1', padding: 12, textAlign: 'left', borderBottom: '1px solid #334155' }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {visibleRows.map((row, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid rgba(51,65,85,0.5)' }}>
                          {currentHeaders.map((h) => (
                            <td key={h} style={{ padding: 12, verticalAlign: 'top' }}>{cleanText(row[h]) || '—'}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
