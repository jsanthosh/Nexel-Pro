import { getDb } from '../db/database';
import { v4 as uuidv4 } from 'uuid';

export interface DocumentRow {
  id: string;
  title: string;
  created_at: string;
  updated_at: string;
  data: string;
}

export interface DocumentMeta {
  id: string;
  title: string;
  created_at: string;
  updated_at: string;
}

export class DocumentRepository {
  create(title: string, data: string): DocumentMeta {
    const db = getDb();
    const id = uuidv4();
    db.prepare(
      'INSERT INTO documents (id, title, data) VALUES (?, ?, ?)'
    ).run(id, title, data);
    return db.prepare(
      'SELECT id, title, created_at, updated_at FROM documents WHERE id = ?'
    ).get(id) as DocumentMeta;
  }

  findById(id: string): DocumentRow | undefined {
    const db = getDb();
    return db.prepare('SELECT * FROM documents WHERE id = ?').get(id) as DocumentRow | undefined;
  }

  list(limit: number, offset: number, search?: string): { documents: DocumentMeta[]; total: number } {
    const db = getDb();
    const params: unknown[] = [];
    let whereClause = '';
    if (search) {
      whereClause = 'WHERE title LIKE ?';
      params.push(`%${search}%`);
    }
    const totalRow = db.prepare(
      `SELECT COUNT(*) as count FROM documents ${whereClause}`
    ).get(...params) as { count: number };
    const documents = db.prepare(
      `SELECT id, title, created_at, updated_at FROM documents ${whereClause} ORDER BY updated_at DESC LIMIT ? OFFSET ?`
    ).all(...params, limit, offset) as DocumentMeta[];
    return { documents, total: totalRow.count };
  }

  update(id: string, title: string, data: string): boolean {
    const db = getDb();
    const result = db.prepare(
      "UPDATE documents SET title = ?, data = ?, updated_at = datetime('now') WHERE id = ?"
    ).run(title, data, id);
    return result.changes > 0;
  }

  updateTitle(id: string, title: string): boolean {
    const db = getDb();
    const result = db.prepare(
      "UPDATE documents SET title = ?, updated_at = datetime('now') WHERE id = ?"
    ).run(title, id);
    return result.changes > 0;
  }

  delete(id: string): boolean {
    const db = getDb();
    const result = db.prepare('DELETE FROM documents WHERE id = ?').run(id);
    return result.changes > 0;
  }
}
