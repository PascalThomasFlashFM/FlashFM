import Database from "better-sqlite3";
import * as fs from "fs";
import * as path from "path";

const SCHEMA = `
CREATE TABLE IF NOT EXISTS sync_state (
  key TEXT PRIMARY KEY,
  value TEXT
);

CREATE TABLE IF NOT EXISTS registrations (
  bocir_id TEXT PRIMARY KEY,
  submission_date INTEGER NOT NULL,
  prenom TEXT NOT NULL,
  nom TEXT NOT NULL,
  email TEXT NOT NULL,
  telephone TEXT,
  date_naissance TEXT,
  question TEXT,
  sheet_row INTEGER,
  added_to_sheet_at TEXT,
  mail_status TEXT NOT NULL DEFAULT 'en_attente',
  exclusion_reason TEXT,
  mail_sent_at TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS sync_runs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  started_at TEXT NOT NULL,
  finished_at TEXT,
  status TEXT NOT NULL,
  nb_processed INTEGER NOT NULL DEFAULT 0,
  nb_errors INTEGER NOT NULL DEFAULT 0,
  error_detail TEXT
);
`;

export type MailStatus = "en_attente" | "envoye" | "exclu";

export interface Registration {
  bocir_id: string;
  submission_date: number;
  prenom: string;
  nom: string;
  email: string;
  telephone: string | null;
  date_naissance: string | null;
  question: string | null;
  sheet_row: number | null;
  added_to_sheet_at: string | null;
  mail_status: MailStatus;
  exclusion_reason: string | null;
  mail_sent_at: string | null;
  created_at: string;
}

export interface SyncRun {
  id: number;
  started_at: string;
  finished_at: string | null;
  status: "running" | "success" | "partial" | "error";
  nb_processed: number;
  nb_errors: number;
  error_detail: string | null;
}

let db: Database.Database | null = null;

export function openDb(dbPath: string): Database.Database {
  if (db) return db;
  fs.mkdirSync(path.dirname(dbPath), { recursive: true });
  db = new Database(dbPath);
  db.pragma("journal_mode = WAL");
  db.exec(SCHEMA);
  return db;
}

function requireDb(): Database.Database {
  if (!db) throw new Error("Base de donnees non initialisee : appelle openDb() d'abord.");
  return db;
}

const LAST_SUBMISSION_DATE_KEY = "dernier_submissionDate_traite";

export function getLastSubmissionDate(): number {
  const row = requireDb()
    .prepare("SELECT value FROM sync_state WHERE key = ?")
    .get(LAST_SUBMISSION_DATE_KEY) as { value: string } | undefined;
  return row ? parseInt(row.value, 10) : 0;
}

export function setLastSubmissionDate(value: number): void {
  requireDb()
    .prepare(
      "INSERT INTO sync_state (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value"
    )
    .run(LAST_SUBMISSION_DATE_KEY, String(value));
}

export function registrationExists(bocirId: string): boolean {
  const row = requireDb()
    .prepare("SELECT 1 FROM registrations WHERE bocir_id = ?")
    .get(bocirId);
  return !!row;
}

export function insertRegistration(reg: {
  bocir_id: string;
  submission_date: number;
  prenom: string;
  nom: string;
  email: string;
  telephone: string | null;
  date_naissance: string | null;
  question: string | null;
  sheet_row: number;
}): void {
  requireDb()
    .prepare(
      `INSERT INTO registrations
        (bocir_id, submission_date, prenom, nom, email, telephone, date_naissance, question, sheet_row, added_to_sheet_at, mail_status)
       VALUES (@bocir_id, @submission_date, @prenom, @nom, @email, @telephone, @date_naissance, @question, @sheet_row, datetime('now'), 'en_attente')`
    )
    .run(reg);
}

export function getRegistration(bocirId: string): Registration | undefined {
  return requireDb()
    .prepare("SELECT * FROM registrations WHERE bocir_id = ?")
    .get(bocirId) as Registration | undefined;
}

export function listRegistrations(status?: MailStatus): Registration[] {
  if (status) {
    return requireDb()
      .prepare("SELECT * FROM registrations WHERE mail_status = ? ORDER BY submission_date DESC")
      .all(status) as Registration[];
  }
  return requireDb()
    .prepare("SELECT * FROM registrations ORDER BY submission_date DESC")
    .all() as Registration[];
}

export function markExcluded(bocirId: string, reason: string | null): void {
  requireDb()
    .prepare(
      "UPDATE registrations SET mail_status = 'exclu', exclusion_reason = ? WHERE bocir_id = ? AND mail_status = 'en_attente'"
    )
    .run(reason, bocirId);
}

export function markSent(bocirId: string): void {
  requireDb()
    .prepare(
      "UPDATE registrations SET mail_status = 'envoye', mail_sent_at = datetime('now') WHERE bocir_id = ?"
    )
    .run(bocirId);
}

export function markSendFailed(bocirId: string): void {
  requireDb()
    .prepare("UPDATE registrations SET mail_status = 'en_attente' WHERE bocir_id = ?")
    .run(bocirId);
}

export function startSyncRun(): number {
  const info = requireDb()
    .prepare("INSERT INTO sync_runs (started_at, status) VALUES (datetime('now'), 'running')")
    .run();
  return Number(info.lastInsertRowid);
}

export function finishSyncRun(
  id: number,
  status: "success" | "partial" | "error",
  nbProcessed: number,
  nbErrors: number,
  errorDetail: string | null
): void {
  requireDb()
    .prepare(
      `UPDATE sync_runs SET finished_at = datetime('now'), status = ?, nb_processed = ?, nb_errors = ?, error_detail = ?
       WHERE id = ?`
    )
    .run(status, nbProcessed, nbErrors, errorDetail, id);
}

export function listSyncRuns(limit = 30): SyncRun[] {
  return requireDb()
    .prepare("SELECT * FROM sync_runs ORDER BY id DESC LIMIT ?")
    .all(limit) as SyncRun[];
}
