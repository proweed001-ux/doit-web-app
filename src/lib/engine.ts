export type Cell = string | number | boolean | null;
export type Row = Record<string, Cell>;
export type Profile = 'balanced' | 'strict' | 'forensic';
export type CanonicalField = 'invoiceNo' | 'invoiceAmount' | 'store' | 'date' | 'time' | 'sku' | 'item' | 'qty' | 'unitPrice' | 'lineTotal' | 'customer' | 'category';
export type Mapping = Partial<Record<CanonicalField, string>>;

export const FIELD_RULES: Record<CanonicalField, RegExp[]> = {
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

export function safeText(value: unknown): string {
  return String(value ?? '').replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim();
}

export function normalize(value: unknown): string {
  return safeText(value)
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[\s\-_\/().,|]+/g, '')
    .replace(/[^\p{L}\p{N}]+/gu, '');
}

export function colIndex(ref: string): number {
  const match = /([A-Z]+)/i.exec(ref);
  if (!match) return 0;
  return match[1].toUpperCase().split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
}

export function colName(index: number): string {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - mod) / 26);
  }
  return s;
}

export function splitCsv(line: string): string[] {
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

export function parseDateParts(text: string): { date: string; time: string; hour: number | null } {
  const raw = safeText(text);
  if (!raw) return { date: '', time: '', hour: null };
  const m = raw.match(/(\d{1,2})[:.](\d{2})(?::(\d{2}))?/);
  if (m) {
    const hour = Number(m[1]);
    return { date: raw, time: `${m[1].padStart(2, '0')}:${m[2]}`, hour: Number.isFinite(hour) ? hour : null };
  }
  const d = new Date(raw);
  if (!Number.isNaN(d.getTime())) {
    return { date: d.toISOString().slice(0, 10), time: d.toTimeString().slice(0, 8), hour: d.getHours() };
  }
  return { date: raw, time: '', hour: null };
}

export function bucketHour(hour: number | null): string {
  if (hour === null || Number.isNaN(hour)) return 'Unknown';
  if (hour < 6) return 'Night';
  if (hour < 12) return 'Morning';
  if (hour < 17) return 'Afternoon';
  if (hour < 21) return 'Evening';
  return 'Late Night';
}

export function makeTemplate(headers: string[]): Mapping {
  const found = (field: CanonicalField): string | undefined => headers.find((h) => FIELD_RULES[field].some((r) => r.test(h)));
  const mapping: Mapping = {};
  (Object.keys(FIELD_RULES) as CanonicalField[]).forEach((field) => {
    mapping[field] = found(field);
  });
  return mapping;
}

export function chooseHeader(headers: string[], value: string | undefined | null): string | undefined {
  if (!value) return undefined;
  return headers.find((h) => h === value) ?? undefined;
}

export function parseSharedStrings(xml: string): string[] {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return Array.from(doc.getElementsByTagName('si')).map((el) => safeText(el.textContent ?? ''));
}

export function parseWorkbookRefs(workbookXml: string, relsXml: string): { name: string; path: string }[] {
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
