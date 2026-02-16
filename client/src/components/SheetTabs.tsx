import React, { useState, useRef, useEffect, useCallback } from 'react';
import { SheetTabConfig } from '../types/spreadsheet';

interface SheetTabsProps {
  tabs: SheetTabConfig[];
  activeSheetId: string;
  onSwitchSheet: (id: string) => void;
  onAddSheet: () => void;
  onDeleteSheet: (id: string) => void;
  onRenameSheet: (id: string, name: string) => void;
  onSetSheetColor: (id: string, color: string | null) => void;
  onReorderSheet: (id: string, newIndex: number) => void;
}

type ContextMenu = { x: number; y: number; tabId: string } | null;

export default function SheetTabs({
  tabs, activeSheetId,
  onSwitchSheet, onAddSheet, onDeleteSheet, onRenameSheet, onSetSheetColor, onReorderSheet,
}: SheetTabsProps) {
  const [contextMenu, setContextMenu] = useState<ContextMenu>(null);
  const [renamingTabId, setRenamingTabId] = useState<string | null>(null);
  const colorInputRef = useRef<HTMLInputElement>(null);
  const colorTabIdRef = useRef<string | null>(null);

  // Drag-to-reorder state
  const [dragTabId, setDragTabId] = useState<string | null>(null);
  const [dropIndex, setDropIndex] = useState<number | null>(null);

  // Close context menu on outside click
  useEffect(() => {
    const close = () => setContextMenu(null);
    document.addEventListener('mousedown', close);
    return () => document.removeEventListener('mousedown', close);
  }, []);

  const commitRename = useCallback((id: string, value: string) => {
    const trimmed = value.trim();
    if (trimmed && trimmed !== tabs.find(t => t.id === id)?.name) {
      onRenameSheet(id, trimmed);
    }
    setRenamingTabId(null);
  }, [tabs, onRenameSheet]);

  const handleContextMenu = useCallback((e: React.MouseEvent, tabId: string) => {
    e.preventDefault();
    e.stopPropagation();
    const menuHeight = 180;
    let y = e.clientY;
    if (y + menuHeight > window.innerHeight) {
      y = Math.max(10, y - menuHeight);
    }
    setContextMenu({ x: e.clientX, y, tabId });
  }, []);

  const startRename = useCallback((tabId: string) => {
    setRenamingTabId(tabId);
    setContextMenu(null);
  }, []);

  const openColorPicker = useCallback((tabId: string) => {
    colorTabIdRef.current = tabId;
    setContextMenu(null);
    setTimeout(() => colorInputRef.current?.click(), 0);
  }, []);

  const moveTab = useCallback((tabId: string, direction: number) => {
    const idx = tabs.findIndex(t => t.id === tabId);
    if (idx === -1) return;
    const newIdx = idx + direction;
    if (newIdx < 0 || newIdx >= tabs.length) return;
    onReorderSheet(tabId, newIdx);
    setContextMenu(null);
  }, [tabs, onReorderSheet]);

  const handleDelete = useCallback((tabId: string) => {
    onDeleteSheet(tabId);
    setContextMenu(null);
  }, [onDeleteSheet]);

  // ── Drag & Drop handlers ──
  const handleDragStart = useCallback((e: React.DragEvent, tabId: string) => {
    setDragTabId(tabId);
    e.dataTransfer.effectAllowed = 'move';
    const el = e.currentTarget as HTMLElement;
    e.dataTransfer.setDragImage(el, el.offsetWidth / 2, el.offsetHeight / 2);
  }, []);

  const handleDragOver = useCallback((e: React.DragEvent, idx: number) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
    const midX = rect.left + rect.width / 2;
    const insertIdx = e.clientX < midX ? idx : idx + 1;
    setDropIndex(insertIdx);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (dragTabId && dropIndex !== null) {
      const fromIdx = tabs.findIndex(t => t.id === dragTabId);
      if (fromIdx !== -1) {
        let targetIdx = dropIndex;
        if (fromIdx < targetIdx) targetIdx--;
        if (targetIdx !== fromIdx && targetIdx >= 0 && targetIdx < tabs.length) {
          onReorderSheet(dragTabId, targetIdx);
        }
      }
    }
    setDragTabId(null);
    setDropIndex(null);
  }, [dragTabId, dropIndex, tabs, onReorderSheet]);

  const handleDragEnd = useCallback(() => {
    setDragTabId(null);
    setDropIndex(null);
  }, []);

  const contextTabIdx = contextMenu ? tabs.findIndex(t => t.id === contextMenu.tabId) : -1;

  return (
    <div className="sheet-tabs-bar">
      <div className="sheet-tabs-scroll">
        {tabs.map((tab, idx) => {
          const isDragging = dragTabId === tab.id;
          const showDropLeft = dropIndex === idx && dragTabId !== null && dragTabId !== tab.id;
          const showDropRight = dropIndex === idx + 1 && dragTabId !== null && dragTabId !== tab.id
            && idx === tabs.length - 1;

          return (
            <div
              key={tab.id}
              className={
                `sheet-tab${tab.id === activeSheetId ? ' sheet-tab--active' : ''}`
                + (isDragging ? ' sheet-tab--dragging' : '')
                + (showDropLeft ? ' sheet-tab--drop-left' : '')
                + (showDropRight ? ' sheet-tab--drop-right' : '')
              }
              style={{
                borderBottomColor: tab.id === activeSheetId
                  ? (tab.color ?? '#1a73e8')
                  : (tab.color ?? 'transparent'),
              }}
              draggable={renamingTabId !== tab.id}
              onClick={() => onSwitchSheet(tab.id)}
              onDoubleClick={() => startRename(tab.id)}
              onContextMenu={(e) => handleContextMenu(e, tab.id)}
              onDragStart={(e) => handleDragStart(e, tab.id)}
              onDragOver={(e) => handleDragOver(e, idx)}
              onDrop={handleDrop}
              onDragEnd={handleDragEnd}
            >
              {renamingTabId === tab.id ? (
                <input
                  className="sheet-tab-rename-input"
                  defaultValue={tab.name}
                  autoFocus
                  onClick={(e) => e.stopPropagation()}
                  onBlur={(e) => commitRename(tab.id, e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') commitRename(tab.id, e.currentTarget.value);
                    if (e.key === 'Escape') setRenamingTabId(null);
                  }}
                />
              ) : (
                <span className="sheet-tab-label">{tab.name}</span>
              )}
            </div>
          );
        })}
      </div>

      <button className="sheet-tab-add-btn" onClick={onAddSheet} title="Add sheet">+</button>

      {/* Hidden color input */}
      <input
        ref={colorInputRef}
        type="color"
        defaultValue="#1a73e8"
        style={{ position: 'absolute', opacity: 0, pointerEvents: 'none' }}
        onChange={(e) => {
          if (colorTabIdRef.current) {
            onSetSheetColor(colorTabIdRef.current, e.target.value);
          }
        }}
      />

      {/* Context menu */}
      {contextMenu && (
        <div
          className="sheet-tab-context-menu"
          style={{ top: contextMenu.y, left: contextMenu.x }}
          onMouseDown={(e) => e.stopPropagation()}
        >
          <button onClick={() => startRename(contextMenu.tabId)}>Rename</button>
          <button onClick={() => handleDelete(contextMenu.tabId)} disabled={tabs.length <= 1}>Delete</button>
          <button onClick={() => openColorPicker(contextMenu.tabId)}>Change Color</button>
          <button onClick={() => moveTab(contextMenu.tabId, -1)} disabled={contextTabIdx <= 0}>Move Left</button>
          <button onClick={() => moveTab(contextMenu.tabId, 1)} disabled={contextTabIdx >= tabs.length - 1}>Move Right</button>
        </div>
      )}
    </div>
  );
}
