import React, { useState } from 'react';
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell as PieCell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from 'recharts';
import { SheetData, CellRange } from '../types/spreadsheet';
import { parseNum } from '../utils/formulaEngine';
import { ChartType } from '../types/chart';

interface ChartPanelProps {
  cells: SheetData;
  selectedRange: CellRange | null;
  chartType?: ChartType;
  onChartTypeChange?: (type: ChartType) => void;
}

const COLORS = ['#1a73e8', '#e8711a', '#34a853', '#ea4335', '#7c3aed', '#f59e0b', '#10b981', '#ef4444'];

function extractChartData(cells: SheetData, range: CellRange | null) {
  if (!range) return { data: [], series: [] };

  const { start, end } = range;
  const rows: string[][] = [];
  for (let r = start.row; r <= end.row; r++) {
    const row: string[] = [];
    for (let c = start.col; c <= end.col; c++) {
      const cell = cells.get(`${r}:${c}`);
      row.push(cell?.displayValue ?? cell?.value ?? '');
    }
    rows.push(row);
  }

  if (rows.length === 0) return { data: [], series: [] };

  const colCount = end.col - start.col + 1;
  // First row = headers if it contains non-numeric values
  const firstRowIsHeader = rows[0].some(v => isNaN(parseNum(v)) && v !== '');
  const headers = firstRowIsHeader ? rows[0] : Array.from({ length: colCount }, (_, i) => `Col ${i + 1}`);
  const dataRows = firstRowIsHeader ? rows.slice(1) : rows;

  // First col = labels, rest = series
  const seriesNames = headers.slice(1);
  const data = dataRows.map(row => {
    const entry: Record<string, string | number> = { label: row[0] ?? '' };
    seriesNames.forEach((name, i) => {
      const val = parseNum(row[i + 1]);
      entry[name] = isNaN(val) ? 0 : val;
    });
    return entry;
  });

  return { data, series: seriesNames };
}

export default function ChartPanel({ cells, selectedRange, chartType: controlledType, onChartTypeChange }: ChartPanelProps) {
  const [internalType, setInternalType] = useState<ChartType>('bar');
  const chartType = controlledType ?? internalType;
  const setChartType = onChartTypeChange ?? setInternalType;
  const { data, series } = extractChartData(cells, selectedRange);

  const empty = data.length === 0 || series.length === 0;

  return (
    <div className="chart-panel">
      <div className="chart-header">
        <span className="chart-title">Chart</span>
        <div className="chart-type-tabs">
          {(['bar', 'line', 'pie'] as ChartType[]).map(t => (
            <button
              key={t}
              className={`chart-tab${chartType === t ? ' chart-tab--active' : ''}`}
              onClick={() => setChartType(t)}
            >
              {t.charAt(0).toUpperCase() + t.slice(1)}
            </button>
          ))}
        </div>
      </div>

      <div className="chart-body">
        {empty ? (
          <div className="chart-empty">Select a data range in the spreadsheet to visualize it.</div>
        ) : chartType === 'bar' ? (
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 24 }}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} />
              <Tooltip />
              <Legend />
              {series.map((s, i) => <Bar key={s} dataKey={s} fill={COLORS[i % COLORS.length]} />)}
            </BarChart>
          </ResponsiveContainer>
        ) : chartType === 'line' ? (
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 24 }}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} />
              <Tooltip />
              <Legend />
              {series.map((s, i) => <Line key={s} type="monotone" dataKey={s} stroke={COLORS[i % COLORS.length]} dot={false} />)}
            </LineChart>
          </ResponsiveContainer>
        ) : (
          <ResponsiveContainer width="100%" height={260}>
            <PieChart>
              <Pie
                data={data.map(d => ({ name: d.label, value: Number(d[series[0]] ?? 0) }))}
                dataKey="value"
                nameKey="name"
                cx="50%"
                cy="50%"
                outerRadius={90}
                label={({ name, percent }) => `${name} ${((percent ?? 0) * 100).toFixed(0)}%`}
              >
                {data.map((_, i) => <PieCell key={i} fill={COLORS[i % COLORS.length]} />)}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        )}
      </div>
    </div>
  );
}
