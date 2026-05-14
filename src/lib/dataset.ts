export type Mapping = Record<string, string | undefined>;
export type Session = { name: string; org: string; token: string; createdAt: string };
export type HistoryItem = { name: string; parsedAt: string; sheets: number; rows: number; anomalies: number };
export type SavedReport = { title: string; generatedAt: string; markdown: string; source: string };
export type DB = { sessions: Session[]; workbooks: HistoryItem[]; reports: SavedReport[]; templates: Record<string, Mapping> };

export const DEFAULT_DB: DB = { sessions: [], workbooks: [], reports: [], templates: {} };

export function downloadJson(filename: string, data: unknown) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function downloadText(filename: string, content: string, type = 'text/plain;charset=utf-8') {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function safeParse<T>(raw: string | null | undefined, fallback: T): T {
  if (!raw) return fallback;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}
