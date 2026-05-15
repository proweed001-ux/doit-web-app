import type { Row } from './engine';
import { normalize, safeText } from './engine';

export function filterRows(rows: Row[], query: string): Row[] {
  if (!query.trim()) return rows;
  const q = normalize(query);
  return rows.filter((row) => Object.entries(row).some(([key, value]) => normalize(key).includes(q) || normalize(value).includes(q)));
}

export function limitRows<T>(rows: T[], limit: number): T[] {
  if (limit <= 0) return [];
  if (rows.length <= limit) return rows;
  return rows.slice(0, limit);
}

export function countRenderableCells(headers: string[], rows: Row[]): number {
  return rows.reduce((sum, row) => sum + headers.reduce((acc, header) => acc + (safeText(row[header]).length > 0 ? 1 : 0), 0), 0);
}

export function summarizeHeaders(headers: string[]): string {
  return headers.slice(0, 8).join(', ');
}
