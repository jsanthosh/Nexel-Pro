import { SerializedDocument } from '../utils/documentSerializer';

export interface DocumentMeta {
  id: string;
  title: string;
  created_at: string;
  updated_at: string;
}

export interface DocumentFull extends DocumentMeta {
  data: SerializedDocument;
}

const BASE = '/api/documents';

export async function createDocument(title?: string): Promise<DocumentMeta> {
  const res = await fetch(BASE, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ title }),
  });
  if (!res.ok) throw new Error(`Create failed: ${res.status}`);
  return res.json();
}

export async function listDocuments(
  limit = 20, offset = 0, search = ''
): Promise<{ documents: DocumentMeta[]; total: number }> {
  const params = new URLSearchParams({ limit: String(limit), offset: String(offset) });
  if (search) params.set('search', search);
  const res = await fetch(`${BASE}?${params}`);
  if (!res.ok) throw new Error(`List failed: ${res.status}`);
  return res.json();
}

export async function loadDocument(id: string): Promise<DocumentFull> {
  const res = await fetch(`${BASE}/${id}`);
  if (!res.ok) throw new Error(`Load failed: ${res.status}`);
  return res.json();
}

export async function saveDocument(id: string, title: string, data: SerializedDocument): Promise<void> {
  const res = await fetch(`${BASE}/${id}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ title, data }),
  });
  if (!res.ok) throw new Error(`Save failed: ${res.status}`);
}

export async function renameDocument(id: string, title: string): Promise<void> {
  const res = await fetch(`${BASE}/${id}`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ title }),
  });
  if (!res.ok) throw new Error(`Rename failed: ${res.status}`);
}

export async function deleteDocument(id: string): Promise<void> {
  const res = await fetch(`${BASE}/${id}`, { method: 'DELETE' });
  if (!res.ok) throw new Error(`Delete failed: ${res.status}`);
}

// --- Versions ---

export interface VersionMeta {
  id: string;
  document_id: string;
  label: string;
  created_at: string;
}

export async function createVersion(
  docId: string,
  label: string,
  data: SerializedDocument
): Promise<VersionMeta> {
  const res = await fetch(`${BASE}/${docId}/versions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ label, data }),
  });
  if (!res.ok) throw new Error(`Create version failed: ${res.status}`);
  return res.json();
}

export async function listVersions(
  docId: string
): Promise<{ versions: VersionMeta[] }> {
  const res = await fetch(`${BASE}/${docId}/versions`);
  if (!res.ok) throw new Error(`List versions failed: ${res.status}`);
  return res.json();
}

export async function restoreVersion(
  docId: string,
  versionId: string
): Promise<{ id: string; data: SerializedDocument }> {
  const res = await fetch(`${BASE}/${docId}/versions/${versionId}/restore`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
  });
  if (!res.ok) throw new Error(`Restore version failed: ${res.status}`);
  return res.json();
}
