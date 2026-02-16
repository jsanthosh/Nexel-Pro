import { useEffect, useRef, useCallback, useState } from 'react';
import { saveDocument } from '../services/documentApi';
import { serializeWorkbook, WorkbookState } from '../utils/documentSerializer';

export type SaveStatus = 'saved' | 'saving' | 'unsaved' | 'error';

export function useAutoSave(
  documentId: string | null,
  title: string,
  workbook: WorkbookState,
  debounceMs = 2000,
) {
  const [saveStatus, setSaveStatus] = useState<SaveStatus>('saved');
  const timeoutRef = useRef<ReturnType<typeof setTimeout>>(undefined);
  const lastSavedRef = useRef<string>('');
  const workbookRef = useRef(workbook);
  const titleRef = useRef(title);
  workbookRef.current = workbook;
  titleRef.current = title;

  const doSave = useCallback(async () => {
    if (!documentId) return;
    const serialized = serializeWorkbook(workbookRef.current);
    const json = JSON.stringify(serialized);
    if (json === lastSavedRef.current) {
      setSaveStatus('saved');
      return;
    }
    setSaveStatus('saving');
    try {
      await saveDocument(documentId, titleRef.current, serialized);
      lastSavedRef.current = json;
      setSaveStatus('saved');
    } catch {
      setSaveStatus('error');
    }
  }, [documentId]);

  // Debounced auto-save on workbook or title changes
  useEffect(() => {
    if (!documentId) return;
    setSaveStatus('unsaved');
    clearTimeout(timeoutRef.current);
    timeoutRef.current = setTimeout(doSave, debounceMs);
    return () => clearTimeout(timeoutRef.current);
  }, [workbook, title, doSave, debounceMs, documentId]);

  // Save on page unload
  useEffect(() => {
    const handler = () => { doSave(); };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [doSave]);

  return { saveStatus, saveNow: doSave };
}
