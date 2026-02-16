import React, { useEffect, useState, useCallback } from 'react';
import { useNavigate } from 'react-router-dom';
import { listDocuments, createDocument, deleteDocument, DocumentMeta } from '../services/documentApi';
import '../App.css';

export default function DocumentList() {
  const navigate = useNavigate();
  const [documents, setDocuments] = useState<DocumentMeta[]>([]);
  const [total, setTotal] = useState(0);
  const [search, setSearch] = useState('');
  const [loading, setLoading] = useState(true);

  const fetchDocs = useCallback(async (query = '') => {
    setLoading(true);
    try {
      const data = await listDocuments(50, 0, query);
      setDocuments(data.documents);
      setTotal(data.total);
    } catch (err) {
      console.error('Failed to load documents', err);
    }
    setLoading(false);
  }, []);

  useEffect(() => { fetchDocs(); }, [fetchDocs]);

  const handleSearch = (e: React.FormEvent) => {
    e.preventDefault();
    fetchDocs(search);
  };

  const handleCreate = async () => {
    const doc = await createDocument();
    navigate(`/doc/${doc.id}`);
  };

  const handleDelete = async (id: string, title: string) => {
    if (!window.confirm(`Delete "${title}"?`)) return;
    await deleteDocument(id);
    fetchDocs(search);
  };

  const formatDate = (dateStr: string) => {
    const d = new Date(dateStr + 'Z');
    return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }) +
      ' ' + d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
  };

  return (
    <div className="doc-list-page">
      <div className="doc-list-container">
        <div className="doc-list-header">
          <h1>Spreadsheets</h1>
          <button className="doc-list-create-btn" onClick={handleCreate}>
            + New Spreadsheet
          </button>
        </div>

        <form className="doc-list-search" onSubmit={handleSearch}>
          <input
            type="text"
            placeholder="Search documents..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </form>

        {loading ? (
          <div className="doc-list-empty">Loading...</div>
        ) : documents.length === 0 ? (
          <div className="doc-list-empty">
            {search ? 'No documents match your search.' : 'No documents yet. Create one to get started!'}
          </div>
        ) : (
          <div className="doc-list-grid">
            {documents.map((doc) => (
              <div
                key={doc.id}
                className="doc-list-card"
                onClick={() => navigate(`/doc/${doc.id}`)}
              >
                <div className="doc-list-card-icon">
                  <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="#217346" strokeWidth="1.5">
                    <rect x="3" y="3" width="18" height="18" rx="2" />
                    <line x1="3" y1="9" x2="21" y2="9" />
                    <line x1="3" y1="15" x2="21" y2="15" />
                    <line x1="9" y1="3" x2="9" y2="21" />
                    <line x1="15" y1="3" x2="15" y2="21" />
                  </svg>
                </div>
                <div className="doc-list-card-info">
                  <div className="doc-list-card-title">{doc.title}</div>
                  <div className="doc-list-card-date">Modified {formatDate(doc.updated_at)}</div>
                </div>
                <button
                  className="doc-list-card-delete"
                  onClick={(e) => { e.stopPropagation(); handleDelete(doc.id, doc.title); }}
                  title="Delete"
                >
                  &times;
                </button>
              </div>
            ))}
          </div>
        )}

        {total > 0 && (
          <div className="doc-list-footer">{total} document{total !== 1 ? 's' : ''}</div>
        )}
      </div>
    </div>
  );
}
