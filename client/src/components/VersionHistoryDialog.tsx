import React, { useState, useEffect, useCallback } from 'react';
import {
  VersionMeta,
  listVersions,
  createVersion,
  restoreVersion,
} from '../services/documentApi';
import { SerializedDocument } from '../utils/documentSerializer';

interface VersionHistoryDialogProps {
  docId: string;
  getCurrentData: () => SerializedDocument;
  onRestored: (data: SerializedDocument) => void;
  onClose: () => void;
}

export default function VersionHistoryDialog({
  docId,
  getCurrentData,
  onRestored,
  onClose,
}: VersionHistoryDialogProps) {
  const [versions, setVersions] = useState<VersionMeta[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [label, setLabel] = useState('');
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    listVersions(docId)
      .then(({ versions }) => setVersions(versions))
      .catch(() => setError('Failed to load versions'))
      .finally(() => setLoading(false));
  }, [docId]);

  const handleSaveVersion = useCallback(async () => {
    if (!label.trim()) return;
    setSaving(true);
    setError(null);
    try {
      const data = getCurrentData();
      const newVersion = await createVersion(docId, label.trim(), data);
      setVersions(prev => [newVersion, ...prev].slice(0, 10));
      setLabel('');
    } catch {
      setError('Failed to save version');
    } finally {
      setSaving(false);
    }
  }, [docId, label, getCurrentData]);

  const handleRestore = useCallback(async (versionId: string) => {
    setSaving(true);
    setError(null);
    try {
      const result = await restoreVersion(docId, versionId);
      const { versions: updated } = await listVersions(docId);
      setVersions(updated);
      onRestored(result.data);
      onClose();
    } catch {
      setError('Failed to restore version');
    } finally {
      setSaving(false);
    }
  }, [docId, onRestored, onClose]);

  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div
        className="modal-dialog"
        onMouseDown={(e) => e.stopPropagation()}
        style={{ maxWidth: 520 }}
      >
        <div className="modal-header">
          <span style={{ fontWeight: 600, fontSize: 14 }}>Version History</span>
          <button className="modal-close-btn" onClick={onClose}>&times;</button>
        </div>

        <div className="modal-body">
          {/* Save new version */}
          <div style={{ marginBottom: 20 }}>
            <div style={{ fontWeight: 600, fontSize: 12, color: '#555', marginBottom: 8 }}>
              Save Current State as Version
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <input
                type="text"
                className="cf-input"
                placeholder="Version label (e.g. Before Q3 pivot)"
                value={label}
                onChange={(e) => setLabel(e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Enter') handleSaveVersion(); }}
                style={{ flex: 1 }}
                maxLength={80}
              />
              <button
                className="modal-btn modal-btn--primary"
                onClick={handleSaveVersion}
                disabled={!label.trim() || saving}
              >
                Save
              </button>
            </div>
            <div style={{ fontSize: 11, color: '#9ca3af', marginTop: 4 }}>
              Up to 10 versions. Oldest is replaced when limit is reached.
            </div>
          </div>

          {/* Version list */}
          <div style={{ fontWeight: 600, fontSize: 12, color: '#555', marginBottom: 8 }}>
            Saved Versions ({versions.length} / 10)
          </div>

          {loading && (
            <div style={{ color: '#9ca3af', fontSize: 13 }}>Loading...</div>
          )}

          {!loading && versions.length === 0 && (
            <div style={{ color: '#9ca3af', fontSize: 13 }}>
              No versions saved yet.
            </div>
          )}

          {!loading && versions.map((v) => (
            <div key={v.id} className="cf-rule-item">
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 600, fontSize: 13, color: '#1d1d1f' }}>
                  {v.label}
                </div>
                <div style={{ fontSize: 11, color: '#9ca3af', marginTop: 2 }}>
                  {new Date(v.created_at + 'Z').toLocaleString()}
                </div>
              </div>
              <button
                className="modal-btn modal-btn--primary"
                style={{ fontSize: 11, padding: '5px 12px' }}
                onClick={() => handleRestore(v.id)}
                disabled={saving}
              >
                Restore
              </button>
            </div>
          ))}

          {error && (
            <div style={{ color: '#ef4444', fontSize: 12, marginTop: 12 }}>
              {error}
            </div>
          )}
        </div>

        <div className="modal-footer">
          <button className="modal-btn" onClick={onClose}>Close</button>
        </div>
      </div>
    </div>
  );
}
