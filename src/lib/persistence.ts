import type { DB } from './dataset';

export function parseDatabaseFile(text: string): DB {
  const parsed = JSON.parse(text) as DB;
  return {
    sessions: Array.isArray(parsed.sessions) ? parsed.sessions : [],
    workbooks: Array.isArray(parsed.workbooks) ? parsed.workbooks : [],
    reports: Array.isArray(parsed.reports) ? parsed.reports : [],
    templates: parsed.templates && typeof parsed.templates === 'object' ? parsed.templates : {},
  };
}

export function exportDatabase(db: DB): string {
  return JSON.stringify(db, null, 2);
}
