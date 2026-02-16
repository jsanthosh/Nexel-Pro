import { getDb } from '../db/database';
import { v4 as uuidv4 } from 'uuid';

export interface VersionRow {
  id: string;
  document_id: string;
  label: string;
  data: string;
  created_at: string;
}

export interface VersionMeta {
  id: string;
  document_id: string;
  label: string;
  created_at: string;
}

export class DocumentVersionRepository {
  create(documentId: string, label: string, data: string): VersionMeta {
    const db = getDb();
    const id = uuidv4();
    db.prepare(
      'INSERT INTO document_versions (id, document_id, label, data) VALUES (?, ?, ?, ?)'
    ).run(id, documentId, label, data);
    return db.prepare(
      'SELECT id, document_id, label, created_at FROM document_versions WHERE id = ?'
    ).get(id) as VersionMeta;
  }

  listByDocument(documentId: string): VersionMeta[] {
    const db = getDb();
    return db.prepare(
      'SELECT id, document_id, label, created_at FROM document_versions WHERE document_id = ? ORDER BY created_at DESC'
    ).all(documentId) as VersionMeta[];
  }

  countByDocument(documentId: string): number {
    const db = getDb();
    const row = db.prepare(
      'SELECT COUNT(*) as count FROM document_versions WHERE document_id = ?'
    ).get(documentId) as { count: number };
    return row.count;
  }

  deleteOldest(documentId: string): void {
    const db = getDb();
    db.prepare(`
      DELETE FROM document_versions WHERE id = (
        SELECT id FROM document_versions
        WHERE document_id = ?
        ORDER BY created_at ASC
        LIMIT 1
      )
    `).run(documentId);
  }

  findById(versionId: string): VersionRow | undefined {
    const db = getDb();
    return db.prepare(
      'SELECT * FROM document_versions WHERE id = ?'
    ).get(versionId) as VersionRow | undefined;
  }
}
