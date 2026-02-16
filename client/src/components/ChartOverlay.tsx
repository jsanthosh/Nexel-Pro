import React, { useRef, useEffect, useCallback } from 'react';
import ChartPanel from './ChartPanel';
import { SheetData, CellRange } from '../types/spreadsheet';
import { ChartType } from '../types/chart';

interface ChartOverlayProps {
  id: string;
  chartType: ChartType;
  dataRange: CellRange | null;
  position: { x: number; y: number };
  size: { width: number; height: number };
  cells: SheetData;
  onPositionChange: (id: string, pos: { x: number; y: number }) => void;
  onChartTypeChange: (id: string, type: ChartType) => void;
  onDelete: (id: string) => void;
  onClone: (id: string) => void;
}

export default function ChartOverlay({
  id, chartType, dataRange, position, size, cells,
  onPositionChange, onChartTypeChange, onDelete, onClone,
}: ChartOverlayProps) {
  const dragRef = useRef<{ startX: number; startY: number; startPos: { x: number; y: number } } | null>(null);

  const handleMove = useCallback((e: MouseEvent) => {
    if (!dragRef.current) return;
    const dx = e.clientX - dragRef.current.startX;
    const dy = e.clientY - dragRef.current.startY;
    onPositionChange(id, {
      x: Math.max(0, dragRef.current.startPos.x + dx),
      y: Math.max(0, dragRef.current.startPos.y + dy),
    });
  }, [id, onPositionChange]);

  const handleUp = useCallback(() => { dragRef.current = null; }, []);

  useEffect(() => {
    document.addEventListener('mousemove', handleMove);
    document.addEventListener('mouseup', handleUp);
    return () => {
      document.removeEventListener('mousemove', handleMove);
      document.removeEventListener('mouseup', handleUp);
    };
  }, [handleMove, handleUp]);

  const handleDragStart = (e: React.MouseEvent) => {
    e.preventDefault();
    dragRef.current = { startX: e.clientX, startY: e.clientY, startPos: position };
  };

  return (
    <div
      className="chart-overlay"
      style={{
        left: position.x,
        top: position.y,
        width: size.width,
        height: size.height,
      }}
      onMouseDown={(e) => e.stopPropagation()}
    >
      <div className="chart-overlay-header" onMouseDown={handleDragStart}>
        <span className="chart-overlay-title">Chart</span>
        <div className="chart-overlay-actions">
          <button onClick={() => onClone(id)} title="Clone chart">Clone</button>
          <button onClick={() => onDelete(id)} title="Delete chart">&times;</button>
        </div>
      </div>
      <ChartPanel
        cells={cells}
        selectedRange={dataRange}
        chartType={chartType}
        onChartTypeChange={(type) => onChartTypeChange(id, type)}
      />
    </div>
  );
}
