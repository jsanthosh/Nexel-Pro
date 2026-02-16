import { Router, Request, Response } from 'express';
import { DocumentRepository } from '../repositories/documentRepository';
import { DocumentVersionRepository } from '../repositories/documentVersionRepository';

const router = Router();
const repo = new DocumentRepository();
const versionRepo = new DocumentVersionRepository();
const MAX_VERSIONS = 10;

const DEFAULT_WORKBOOK = JSON.stringify({
  version: 1,
  sheets: {
    'sheet-default': {
      cells: {},
      rowCount: 100,
      colCount: 26,
      colWidths: {},
      freezeRows: 0,
      conditionalFormatRules: [],
      tables: [],
    },
  },
  tabs: [{ id: 'sheet-default', name: 'Sheet1', color: null }],
  activeSheetId: 'sheet-default',
});

// Create new document
router.post('/documents', (req: Request, res: Response) => {
  const title = req.body.title || 'Untitled Spreadsheet';
  const meta = repo.create(title, DEFAULT_WORKBOOK);
  res.status(201).json(meta);
});

// List documents
router.get('/documents', (req: Request, res: Response) => {
  const limit = Math.min(Number(req.query.limit) || 20, 100);
  const offset = Number(req.query.offset) || 0;
  const search = (req.query.search as string) || '';
  res.json(repo.list(limit, offset, search || undefined));
});

// Load document
router.get('/documents/:id', (req: Request, res: Response) => {
  const doc = repo.findById(req.params.id);
  if (!doc) { res.status(404).json({ error: 'Document not found' }); return; }
  res.json({ id: doc.id, title: doc.title, created_at: doc.created_at, updated_at: doc.updated_at, data: JSON.parse(doc.data) });
});

// Save document (full update)
router.put('/documents/:id', (req: Request, res: Response) => {
  const { title, data } = req.body;
  if (!data) { res.status(400).json({ error: 'data is required' }); return; }
  const success = repo.update(req.params.id, title || 'Untitled Spreadsheet', JSON.stringify(data));
  if (!success) { res.status(404).json({ error: 'Document not found' }); return; }
  res.json({ id: req.params.id, title, updated_at: new Date().toISOString() });
});

// Rename document
router.patch('/documents/:id', (req: Request, res: Response) => {
  const { title } = req.body;
  if (!title) { res.status(400).json({ error: 'title is required' }); return; }
  const success = repo.updateTitle(req.params.id, title);
  if (!success) { res.status(404).json({ error: 'Document not found' }); return; }
  res.json({ id: req.params.id, title, updated_at: new Date().toISOString() });
});

// Delete document
router.delete('/documents/:id', (req: Request, res: Response) => {
  repo.delete(req.params.id);
  res.status(204).send();
});

// Save current state as a named version
router.post('/documents/:id/versions', (req: Request, res: Response) => {
  const doc = repo.findById(req.params.id);
  if (!doc) { res.status(404).json({ error: 'Document not found' }); return; }

  const label = (req.body.label as string) || `Version ${new Date().toLocaleString()}`;
  const data = req.body.data ? JSON.stringify(req.body.data) : doc.data;

  const count = versionRepo.countByDocument(req.params.id);
  if (count >= MAX_VERSIONS) {
    versionRepo.deleteOldest(req.params.id);
  }

  const version = versionRepo.create(req.params.id, label, data);
  res.status(201).json(version);
});

// List versions (metadata only)
router.get('/documents/:id/versions', (req: Request, res: Response) => {
  const doc = repo.findById(req.params.id);
  if (!doc) { res.status(404).json({ error: 'Document not found' }); return; }
  const versions = versionRepo.listByDocument(req.params.id);
  res.json({ versions });
});

// Restore a version (auto-saves current state first)
router.post('/documents/:id/versions/:versionId/restore', (req: Request, res: Response) => {
  const doc = repo.findById(req.params.id);
  if (!doc) { res.status(404).json({ error: 'Document not found' }); return; }

  const targetVersion = versionRepo.findById(req.params.versionId);
  if (!targetVersion || targetVersion.document_id !== req.params.id) {
    res.status(404).json({ error: 'Version not found' });
    return;
  }

  // Auto-save current state before restoring
  const autoLabel = `Before restore — ${new Date().toISOString().slice(0, 19).replace('T', ' ')}`;
  const count = versionRepo.countByDocument(req.params.id);
  if (count >= MAX_VERSIONS) {
    versionRepo.deleteOldest(req.params.id);
  }
  versionRepo.create(req.params.id, autoLabel, doc.data);

  // Replace document data with the version's data
  repo.update(req.params.id, doc.title, targetVersion.data);

  res.json({ id: req.params.id, data: JSON.parse(targetVersion.data) });
});

export default router;
