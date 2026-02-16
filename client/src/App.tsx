import React, { useState, useCallback, useEffect, lazy, Suspense } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { useSpreadsheet } from './hooks/useSpreadsheet';
import { useChat } from './hooks/useChat';
import { useAutoSave, SaveStatus } from './hooks/useAutoSave';
import { loadDocument, createDocument, renameDocument as renameDocApi } from './services/documentApi';
import { deserializeWorkbook, serializeWorkbook } from './utils/documentSerializer';
import SpreadsheetGrid from './components/SpreadsheetGrid';
import Toolbar from './components/Toolbar';
import FormulaBar from './components/FormulaBar';
import ChartOverlay from './components/ChartOverlay';
import SheetTabs from './components/SheetTabs';
import ConditionalFormatDialog from './components/ConditionalFormatDialog';
import PivotTableDialog from './components/PivotTableDialog';
import VersionHistoryDialog from './components/VersionHistoryDialog';
import DocumentTitle from './components/DocumentTitle';
import { CellRange } from './types/spreadsheet';
import { ChartInstance, ChartType } from './types/chart';
import './App.css';

const ChatPanel = lazy(() => import('./components/ChatPanel'));

const STATUS_LABELS: Record<SaveStatus, string> = {
  saved: 'Saved',
  saving: 'Saving...',
  unsaved: 'Unsaved changes',
  error: 'Save failed',
};

export default function App() {
  const { id: routeId } = useParams<{ id: string }>();
  const navigate = useNavigate();

  const [docId, setDocId] = useState<string | null>(null);
  const [docTitle, setDocTitle] = useState('Untitled Spreadsheet');
  const [loading, setLoading] = useState(true);

  const spreadsheet = useSpreadsheet();
  const { saveStatus, saveNow } = useAutoSave(docId, docTitle, spreadsheet._workbook);

  // Load or create document on mount
  useEffect(() => {
    let cancelled = false;

    async function init() {
      try {
        if (routeId && routeId !== 'new') {
          // Load existing document
          const doc = await loadDocument(routeId);
          if (cancelled) return;
          setDocId(doc.id);
          setDocTitle(doc.title);
          spreadsheet.loadWorkbook(deserializeWorkbook(doc.data));
        } else {
          // Create new document
          const doc = await createDocument();
          if (cancelled) return;
          setDocId(doc.id);
          setDocTitle(doc.title);
          navigate(`/doc/${doc.id}`, { replace: true });
        }
      } catch (err) {
        console.error('Failed to load/create document', err);
        if (!cancelled) navigate('/', { replace: true });
      }
      if (!cancelled) setLoading(false);
    }

    init();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [routeId]);

  const handleRename = useCallback(async (newTitle: string) => {
    setDocTitle(newTitle);
    if (docId) {
      try { await renameDocApi(docId, newTitle); } catch { /* auto-save will persist it */ }
    }
  }, [docId]);

  const [chartInstances, setChartInstances] = useState<ChartInstance[]>([]);
  const [showCFDialog, setShowCFDialog] = useState(false);
  const [showPivotDialog, setShowPivotDialog] = useState(false);
  const [showVersionDialog, setShowVersionDialog] = useState(false);

  const getCurrentData = useCallback(
    () => serializeWorkbook(spreadsheet._workbook),
    [spreadsheet._workbook]
  );

  const handleVersionRestored = useCallback((data: any) => {
    spreadsheet.loadWorkbook(deserializeWorkbook(data));
  }, [spreadsheet]);

  const handleAddChart = useCallback(() => {
    const newChart: ChartInstance = {
      id: `chart-${Date.now()}`,
      chartType: 'bar',
      dataRange: spreadsheet.selectedRange,
      position: { x: 50 + chartInstances.length * 30, y: 50 + chartInstances.length * 30 },
      size: { width: 500, height: 340 },
    };
    setChartInstances(prev => [...prev, newChart]);
  }, [spreadsheet.selectedRange, chartInstances.length]);

  const handleAddChartFromAI = useCallback((dataRange: CellRange, chartType?: ChartType) => {
    const newChart: ChartInstance = {
      id: `chart-${Date.now()}`,
      chartType: chartType ?? 'bar',
      dataRange,
      position: { x: 50 + chartInstances.length * 30, y: 50 + chartInstances.length * 30 },
      size: { width: 500, height: 340 },
    };
    setChartInstances(prev => [...prev, newChart]);
  }, [chartInstances.length]);

  const chat = useChat(spreadsheet, handleAddChartFromAI);

  const handleDeleteChart = useCallback((id: string) => {
    setChartInstances(prev => prev.filter(c => c.id !== id));
  }, []);

  const handleCloneChart = useCallback((id: string) => {
    setChartInstances(prev => {
      const original = prev.find(c => c.id === id);
      if (!original) return prev;
      return [...prev, {
        ...original,
        id: `chart-${Date.now()}`,
        position: { x: original.position.x + 30, y: original.position.y + 30 },
      }];
    });
  }, []);

  const handleChartPositionChange = useCallback((id: string, pos: { x: number; y: number }) => {
    setChartInstances(prev => prev.map(c => c.id === id ? { ...c, position: pos } : c));
  }, []);

  const handleChartTypeChange = useCallback((id: string, chartType: ChartType) => {
    setChartInstances(prev => prev.map(c => c.id === id ? { ...c, chartType } : c));
  }, []);

  if (loading) {
    return (
      <div className="app">
        <div className="app-loading">Loading spreadsheet...</div>
      </div>
    );
  }

  return (
    <div className="app">
      <div className="app-header">
        <div className="app-header-left">
          <button className="app-back-btn" onClick={() => { saveNow(); navigate('/'); }} title="All documents">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
          </button>
          <DocumentTitle title={docTitle} onRename={handleRename} />
          <span className={`save-status save-status--${saveStatus}`}>{STATUS_LABELS[saveStatus]}</span>
          <button
            className="toolbar-btn"
            onClick={() => setShowVersionDialog(true)}
            title="Version History"
            style={{ fontSize: 12, marginLeft: 4 }}
          >
            History
          </button>
        </div>
        <Toolbar
          selectedRange={spreadsheet.selectedRange}
          activeFormatting={spreadsheet.getActiveCellFormatting()}
          canUndo={spreadsheet.canUndo}
          canRedo={spreadsheet.canRedo}
          onUndo={spreadsheet.undo}
          onRedo={spreadsheet.redo}
          onFormat={spreadsheet.applyFormat}
          onColor={spreadsheet.applyColor}
          onExport={spreadsheet.exportCSV}
          onImport={spreadsheet.importCSV}
          onAddChart={handleAddChart}
          onOpenConditionalFormat={() => setShowCFDialog(true)}
          onOpenPivotTable={() => setShowPivotDialog(true)}
          onCreateTable={spreadsheet.createTable}
          tables={spreadsheet.tables}
          onSetTableStyle={spreadsheet.setTableStyle}
        />
      </div>

      <div className="app-body">
        <div className="spreadsheet-area">
          <FormulaBar
            activeCell={spreadsheet.activeCell}
            value={spreadsheet.getActiveCellValue()}
            onChange={spreadsheet.setCellValue}
          />
          <SpreadsheetGrid
            cells={spreadsheet.cells}
            rowCount={spreadsheet.rowCount}
            colCount={spreadsheet.colCount}
            selectedRange={spreadsheet.selectedRange}
            activeCell={spreadsheet.activeCell}
            colWidths={spreadsheet.colWidths}
            rowHeights={spreadsheet.rowHeights}
            freezeRows={spreadsheet.freezeRows}
            conditionalFormatRules={spreadsheet.conditionalFormatRules}
            tables={spreadsheet.tables}
            onCellChange={spreadsheet.setCellValue}
            onSelectionChange={spreadsheet.setSelection}
            onUndo={spreadsheet.undo}
            onRedo={spreadsheet.redo}
            onFillDown={spreadsheet.fillDown}
            onFillRight={spreadsheet.fillRight}
            onFillSeries={spreadsheet.fillSeries}
            onColWidthChange={spreadsheet.setColWidth}
            onRowHeightChange={spreadsheet.setRowHeight}
            onInsertRow={spreadsheet.insertRow}
            onDeleteRow={spreadsheet.deleteRow}
            onInsertCol={spreadsheet.insertCol}
            onDeleteCol={spreadsheet.deleteCol}
            onSetFreezeRows={spreadsheet.setFreezeRows}
            onSortTable={spreadsheet.sortTable}
            onFilterTable={spreadsheet.filterTable}
            onDeleteTable={spreadsheet.deleteTable}
            onFormat={spreadsheet.applyFormat}
          />
          {chartInstances.map(chart => (
            <ChartOverlay
              key={chart.id}
              id={chart.id}
              chartType={chart.chartType}
              dataRange={chart.dataRange}
              position={chart.position}
              size={chart.size}
              cells={spreadsheet.cells}
              onPositionChange={handleChartPositionChange}
              onChartTypeChange={handleChartTypeChange}
              onDelete={handleDeleteChart}
              onClone={handleCloneChart}
            />
          ))}
          <SheetTabs
            tabs={spreadsheet.tabs}
            activeSheetId={spreadsheet.activeSheetId}
            onSwitchSheet={spreadsheet.switchSheet}
            onAddSheet={spreadsheet.addSheet}
            onDeleteSheet={spreadsheet.deleteSheet}
            onRenameSheet={spreadsheet.renameSheet}
            onSetSheetColor={spreadsheet.setSheetColor}
            onReorderSheet={spreadsheet.reorderSheet}
          />
        </div>

        <Suspense fallback={<div className="chat-panel">Loading chat...</div>}>
          <ChatPanel
            messages={chat.messages}
            isLoading={chat.isLoading}
            onSend={chat.sendMessage}
            model={chat.model}
            onModelChange={chat.setModel}
          />
        </Suspense>
      </div>

      {showCFDialog && (
        <ConditionalFormatDialog
          rules={spreadsheet.conditionalFormatRules}
          selectedRange={spreadsheet.selectedRange}
          onAdd={spreadsheet.addCFRule}
          onDelete={spreadsheet.deleteCFRule}
          onClose={() => setShowCFDialog(false)}
        />
      )}

      {showPivotDialog && (
        <PivotTableDialog
          cells={spreadsheet.cells}
          selectedRange={spreadsheet.selectedRange}
          onClose={() => setShowPivotDialog(false)}
        />
      )}

      {showVersionDialog && docId && (
        <VersionHistoryDialog
          docId={docId}
          getCurrentData={getCurrentData}
          onRestored={handleVersionRestored}
          onClose={() => setShowVersionDialog(false)}
        />
      )}
    </div>
  );
}
