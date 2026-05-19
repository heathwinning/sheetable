import React, { useState, useCallback, useEffect, useRef } from 'react';
import { useParams, useNavigate, useSearchParams } from 'react-router-dom';
import Select from 'react-select';
import { dialogSelectStyles } from './selectStyles';
import GridLayout, { WidthProvider } from 'react-grid-layout/legacy';
import 'react-grid-layout/css/styles.css';
import 'react-resizable/css/styles.css';
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area, PieChart, Pie, Cell,
  ScatterChart, Scatter, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer,
} from 'recharts';
import { AgGridReact } from 'ag-grid-react';
import { AllCommunityModule, themeQuartz } from 'ag-grid-community';
import type { ColDef, ColGroupDef } from 'ag-grid-community';
import type { UseAppStateReturn } from './useAppState';
import type { ChartConfig, ChartLayoutItem, ChartType, AggregateFunc, Row, DateFeature } from './types';
import { applyChartValueFormat } from './chartFormat';
import { useConfirm } from './DialogProvider';

const RGL = WidthProvider(GridLayout);
const CHART_COLORS = ['#6366f1', '#22d3ee', '#f59e0b', '#10b981', '#ef4444', '#8b5cf6', '#f97316', '#06b6d4'];

const bookPrefix = (bookName?: string) => (bookName ? `/book/${encodeURIComponent(bookName)}` : '');

// ── Aggregation ──────────────────────────────────────────────────────────────

function toNum(v: unknown): number {
  const n = Number(v);
  return Number.isNaN(n) ? 0 : n;
}

function applyAgg(vals: number[], agg: AggregateFunc): number {
  if (vals.length === 0) return 0;
  switch (agg) {
    case 'count': return vals.length;
    case 'sum': return vals.reduce((a, b) => a + b, 0);
    case 'avg': return vals.reduce((a, b) => a + b, 0) / vals.length;
    case 'min': return Math.min(...vals);
    case 'max': return Math.max(...vals);
    default: return vals[0];
  }
}

// Strip :feature suffix from a column expression
const VALID_DATE_FEATURES: DateFeature[] = ['year', 'quarter', 'yearmonth', 'month', 'monthnum', 'week', 'dayofweek', 'day', 'hour'];
function stripFeature(col: string): string {
  const i = col.lastIndexOf(':');
  if (i < 0) return col;
  return VALID_DATE_FEATURES.includes(col.slice(i + 1) as DateFeature) ? col.slice(0, i) : col;
}
function getFeature(col: string): DateFeature | undefined {
  const i = col.lastIndexOf(':');
  if (i < 0) return undefined;
  const f = col.slice(i + 1) as DateFeature;
  return VALID_DATE_FEATURES.includes(f) ? f : undefined;
}

// Format a numeric value using chartFormat (calc + template), falling back to legacy ColumnModifier
function formatValue(n: number, config: Pick<ChartConfig, 'valueFormat' | 'yModifier'>): string {
  return applyChartValueFormat(n, config);
}

function aggregateData(
  rows: Row[],
  xCol: string,
  yCol: string,
  agg: AggregateFunc,
  groupBy: string | undefined,
  resolveColumnPath: (tableName: string, row: Row, path: string) => string,
  tableName: string,
): { data: Record<string, unknown>[]; seriesKeys: string[] } {
  if (!xCol) return { data: [], seriesKeys: [] };

  // Parse "colpath:feature" expressions — colpath may be a dot-path like "species.name"
  const parseExpr = (expr: string): { col: string; feature: DateFeature | null } => {
    // Find last colon that looks like a date feature (not a dot-path segment)
    const colon = expr.lastIndexOf(':');
    if (colon < 0) return { col: expr, feature: null };
    const maybeFeature = expr.slice(colon + 1) as DateFeature;
    const validFeatures: DateFeature[] = ['year', 'quarter', 'yearmonth', 'month', 'monthnum', 'week', 'dayofweek', 'day', 'hour'];
    if (!validFeatures.includes(maybeFeature)) return { col: expr, feature: null };
    return { col: expr.slice(0, colon), feature: maybeFeature };
  };

  const xExpr = parseExpr(xCol);
  const gExpr = groupBy ? parseExpr(groupBy) : null;

  const extractValue = (row: Row, expr: { col: string; feature: DateFeature | null }): string => {
    const raw = resolveColumnPath(tableName, row, expr.col);
    if (!raw) return '';
    if (!expr.feature) return raw;
    // Extract date feature
    const d = new Date(raw);
    if (isNaN(d.getTime())) return raw;
    switch (expr.feature) {
      case 'year': return String(d.getFullYear());
      case 'quarter': return `Q${Math.ceil((d.getMonth() + 1) / 3)}`;
      case 'yearmonth': return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`;
      case 'month': return d.toLocaleString('default', { month: 'long' });
      case 'monthnum': return String(d.getMonth() + 1).padStart(2, '0');
      case 'week': {
        const jan1 = new Date(d.getFullYear(), 0, 1);
        return String(Math.ceil(((d.getTime() - jan1.getTime()) / 86400000 + jan1.getDay() + 1) / 7)).padStart(2, '0');
      }
      case 'dayofweek': return d.toLocaleString('default', { weekday: 'long' });
      case 'day': return String(d.getDate()).padStart(2, '0');
      case 'hour': return String(d.getHours()).padStart(2, '0') + ':00';
      default: return raw;
    }
  };

  if (groupBy) {
    const xOrder: string[] = [];
    const seenX = new Set<string>();
    const seenG = new Set<string>();
    const groups = new Map<string, Map<string, number[]>>();
    for (const row of rows) {
      const x = extractValue(row, xExpr);
      const g = extractValue(row, gExpr!);
      if (!seenX.has(x)) { xOrder.push(x); seenX.add(x); }
      seenG.add(g);
      if (!groups.has(x)) groups.set(x, new Map());
      const xg = groups.get(x)!;
      if (!xg.has(g)) xg.set(g, []);
      xg.get(g)!.push(agg === 'count' ? 1 : toNum(resolveColumnPath(tableName, row, yCol)));
    }
    const seriesKeys = Array.from(seenG);
    const data = xOrder.map(x => {
      const xg = groups.get(x) ?? new Map();
      const entry: Record<string, unknown> = { x };
      for (const g of seriesKeys) entry[g] = applyAgg(xg.get(g) ?? [], agg);
      return entry;
    });
    return { data, seriesKeys };
  }

  const xOrder: string[] = [];
  const seenX = new Set<string>();
  const groups = new Map<string, number[]>();
  for (const row of rows) {
    const x = extractValue(row, xExpr);
    if (!seenX.has(x)) { xOrder.push(x); seenX.add(x); }
    if (!groups.has(x)) groups.set(x, []);
    groups.get(x)!.push(agg === 'count' ? 1 : toNum(resolveColumnPath(tableName, row, yCol)));
  }
  const seriesKey = yCol || 'value';
  const data = xOrder.map(x => ({ x, [seriesKey]: applyAgg(groups.get(x) ?? [], agg) }));
  return { data, seriesKeys: [seriesKey] };
}

// ── Pivot table data builder ─────────────────────────────────────────────────

interface PivotResult {
  rowDims: string[];        // display labels for row dimension headers
  colDims: string[];        // display labels for column dimension headers
  rowKeys: string[][];      // unique combos of row dimension values, in order
  colKeys: string[][];      // unique combos of col dimension values, in order (empty if no colDims)
  cells: Map<string, number>; // `rowJoined\1colJoined` -> aggregated value
  rowTotals: Map<string, number>; // rowJoined -> row total
  colTotals: Map<string, number>; // colJoined -> col total
  grandTotal: number;
}

function buildPivotTable(
  rows: Row[],
  rowDims: string[],
  colDims: string[],
  yCol: string,
  agg: AggregateFunc,
  resolveColumnPath: (tableName: string, row: Row, path: string) => string,
  tableName: string,
): PivotResult {
  const parseExpr = (expr: string): { col: string; feature: DateFeature | null } => {
    const colon = expr.lastIndexOf(':');
    if (colon < 0) return { col: expr, feature: null };
    const maybeFeature = expr.slice(colon + 1) as DateFeature;
    const validFeatures: DateFeature[] = ['year', 'quarter', 'yearmonth', 'month', 'monthnum', 'week', 'dayofweek', 'day', 'hour'];
    if (!validFeatures.includes(maybeFeature)) return { col: expr, feature: null };
    return { col: expr.slice(0, colon), feature: maybeFeature };
  };

  const extractVal = (row: Row, expr: { col: string; feature: DateFeature | null }): string => {
    const raw = resolveColumnPath(tableName, row, expr.col);
    if (!raw) return '';
    // For reference columns with no date feature, use only the first display column value
    if (!expr.feature) return raw.includes(' · ') ? raw.split(' · ')[0] : raw;
    const d = new Date(raw);
    if (isNaN(d.getTime())) return raw;
    switch (expr.feature) {
      case 'year': return String(d.getFullYear());
      case 'quarter': return `Q${Math.ceil((d.getMonth() + 1) / 3)}`;
      case 'yearmonth': return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`;
      case 'month': return d.toLocaleString('default', { month: 'long' });
      case 'monthnum': return String(d.getMonth() + 1).padStart(2, '0');
      case 'week': { const j = new Date(d.getFullYear(), 0, 1); return String(Math.ceil(((d.getTime() - j.getTime()) / 86400000 + j.getDay() + 1) / 7)).padStart(2, '0'); }
      case 'dayofweek': return d.toLocaleString('default', { weekday: 'long' });
      case 'day': return String(d.getDate()).padStart(2, '0');
      case 'hour': return String(d.getHours()).padStart(2, '0') + ':00';
      default: return raw;
    }
  };

  const rowExprs = rowDims.map(parseExpr);
  const colExprs = colDims.map(parseExpr);

  const rowKeyOrder: string[] = [];
  const rowKeyMap = new Map<string, string[]>();
  const colKeyOrder: string[] = [];
  const colKeyMap = new Map<string, string[]>();
  // raw numeric values per cell / row / col / all — for correct aggregation of totals
  const cellRaw = new Map<string, number[]>();
  const rowRaw = new Map<string, number[]>();
  const colRaw = new Map<string, number[]>();
  const allRaw: number[] = [];

  for (const row of rows) {
    const rowVals = rowExprs.map(e => extractVal(row, e));
    const colVals = colExprs.map(e => extractVal(row, e));
    const rowJoined = rowVals.join('\0');
    const colJoined = colVals.join('\0');

    if (!rowKeyMap.has(rowJoined)) { rowKeyOrder.push(rowJoined); rowKeyMap.set(rowJoined, rowVals); }
    if (colExprs.length > 0 && !colKeyMap.has(colJoined)) { colKeyOrder.push(colJoined); colKeyMap.set(colJoined, colVals); }

    const yRaw = yCol ? resolveColumnPath(tableName, row, yCol) : '';
    const yNum = agg === 'count' ? 1 : toNum(yRaw);

    const cellKey = `${rowJoined}\x01${colJoined}`;
    if (!cellRaw.has(cellKey)) cellRaw.set(cellKey, []);
    cellRaw.get(cellKey)!.push(yNum);

    if (!rowRaw.has(rowJoined)) rowRaw.set(rowJoined, []);
    rowRaw.get(rowJoined)!.push(yNum);

    if (colExprs.length > 0) {
      if (!colRaw.has(colJoined)) colRaw.set(colJoined, []);
      colRaw.get(colJoined)!.push(yNum);
    }
    allRaw.push(yNum);
  }

  const cells = new Map<string, number>();
  for (const [k, vals] of cellRaw) cells.set(k, applyAgg(vals, agg));
  const rowTotals = new Map<string, number>();
  for (const [k, vals] of rowRaw) rowTotals.set(k, applyAgg(vals, agg));
  const colTotals = new Map<string, number>();
  for (const [k, vals] of colRaw) colTotals.set(k, applyAgg(vals, agg));
  const grandTotal = applyAgg(allRaw, agg);

  // Derive display labels for dimension headers (strip :feature suffix for header)
  const dimLabel = (d: string) => d.includes(':') ? d.split(':')[0] : d;

  return {
    rowDims: rowDims.map(dimLabel),
    colDims: colDims.map(dimLabel),
    rowKeys: rowKeyOrder.map(k => rowKeyMap.get(k)!),
    colKeys: colKeyOrder.map(k => colKeyMap.get(k)!),
    cells,
    rowTotals,
    colTotals,
    grandTotal,
  };
}

// helper: format a pivot display value (handles yearmonth -> MMM YYYY)
function fmtDimVal(val: string, dim: string): string {
  if (dim.endsWith(':yearmonth')) {
    const d = new Date(val);
    if (!isNaN(d.getTime())) return d.toLocaleString('default', { month: 'short', year: 'numeric' });
  }
  return val || '—';
}

// ── Chart renderer ───────────────────────────────────────────────────────────

const ChartRenderer: React.FC<{
  config: ChartConfig;
  data: Record<string, unknown>[];
  seriesKeys: string[];
  rows?: Row[];
  resolveColumnPath?: (tableName: string, row: Row, path: string) => string;
  getColumnPaths?: (tableId: string) => { path: string; label: string; type?: string }[];
  onSortChange?: (sort: { key: string; dir: 'asc' | 'desc' } | undefined) => void;
}> = ({ config, data, seriesKeys, rows, resolveColumnPath, getColumnPaths }) => {
  if (config.type === 'table') {
    const rowDims = config.tableRows?.length ? config.tableRows : (config.xColumn ? [config.xColumn] : []);
    const colDims = config.tableColumns?.length ? config.tableColumns : (config.groupBy ? [config.groupBy] : []);
    if (!rowDims.length || !rows || !resolveColumnPath) {
      return (
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%', color: 'var(--color-text-muted)', fontSize: 13 }}>
          No data
        </div>
      );
    }
    const pivot = buildPivotTable(rows, rowDims, colDims, config.yColumn, config.aggregate, resolveColumnPath, config.table);
    const hasColDims = pivot.colKeys.length > 0;

    // Build label lookup for column paths (strips :feature suffix before lookup)
    const pathLabels = getColumnPaths ? new Map(getColumnPaths(config.table).map(p => [p.path, p.label])) : new Map<string, string>();
    const dimLabel = (dim: string) => {
      const path = stripFeature(dim);
      const feature = getFeature(dim);
      const fullLabel = pathLabels.get(path) ?? (path.includes(':') ? path.split(':')[0] : path);
      // Use only the last segment of the reference chain for brevity
      const base = fullLabel.includes(' → ') ? fullLabel.split(' → ').pop()! : fullLabel;
      const feat = feature ? DATE_FEATURES.find(f => f.value === feature)?.label : undefined;
      return feat ? `${base} (${feat})` : base;
    };

    const valueFormatter = (p: { value: number | null | undefined }) =>
      p.value !== null && p.value !== undefined ? formatValue(p.value, config) : '';

    // Precompute subtotal fields for each group prefix at non-leaf levels
    // field key: `_subtotal_<prefix joined with \x01>`; value: array of _col_N indices in that group
    const subtotalMap = new Map<string, number[]>();
    if (hasColDims && colDims.length > 1) {
      const prefixSeen = new Set<string>();
      for (const ck of pivot.colKeys) {
        for (let lvl = 0; lvl < colDims.length - 1; lvl++) {
          const prefix = ck.slice(0, lvl + 1);
          const stKey = `_subtotal_${prefix.join('\x01')}`;
          if (!prefixSeen.has(stKey)) {
            prefixSeen.add(stKey);
            subtotalMap.set(stKey, pivot.colKeys
              .map((ck2, ci2) => prefix.every((v, i) => ck2[i] === v) ? ci2 : -1)
              .filter(i => i >= 0));
          }
        }
      }
    }

    // Build AG Grid rowData
    const rowData = pivot.rowKeys.map(rk => {
      const rowJoined = rk.join('\0');
      const row: Record<string, unknown> = {};
      rk.forEach((v, i) => { row[`_dim_${i}`] = fmtDimVal(v, rowDims[i]); });
      if (hasColDims) {
        pivot.colKeys.forEach((ck, ci) => {
          const colJoined = ck.join('\0');
          const val = pivot.cells.get(`${rowJoined}\x01${colJoined}`);
          row[`_col_${ci}`] = val ?? null;
        });
        for (const [stKey, indices] of subtotalMap) {
          row[stKey] = indices.reduce((s, ci) => { const v = row[`_col_${ci}`]; return s + (typeof v === 'number' ? v : 0); }, 0);
        }
      } else {
        row['_val'] = pivot.cells.get(`${rowJoined}\x01`) ?? pivot.rowTotals.get(rowJoined) ?? 0;
      }
      row['_total'] = pivot.rowTotals.get(rowJoined) ?? 0;
      return row;
    });

    // Pinned bottom totals row
    const totalRow: Record<string, unknown> = { _dim_0: 'Total' };
    rowDims.slice(1).forEach((_, i) => { totalRow[`_dim_${i + 1}`] = ''; });
    if (hasColDims) {
      pivot.colKeys.forEach((ck, ci) => {
        totalRow[`_col_${ci}`] = pivot.colTotals.get(ck.join('\0')) ?? null;
      });
      for (const [stKey, indices] of subtotalMap) {
        totalRow[stKey] = indices.reduce((s, ci) => { const v = totalRow[`_col_${ci}`]; return s + (typeof v === 'number' ? v : 0); }, 0);
      }
    } else {
      totalRow['_val'] = pivot.grandTotal;
    }
    totalRow['_total'] = pivot.grandTotal;

    // Recursively build nested column defs for multi-level column dimensions
    // Adds a Subtotal column at the end of each group when colDims.length > 1
    const buildColGroup = (keys: string[][], level: number, prefix: string[]): (ColDef | ColGroupDef)[] => {
      const filtered = keys.filter(ck => prefix.every((v, i) => ck[i] === v));
      if (level === colDims.length - 1) {
        return filtered.map(ck => ({
          headerName: fmtDimVal(ck[level], colDims[level]),
          field: `_col_${pivot.colKeys.indexOf(ck)}`,
          sortable: true,
          type: 'numericColumn',
          valueFormatter,
        } as ColDef));
      }
      const seen = new Set<string>();
      const groups: ColGroupDef[] = [];
      for (const ck of filtered) {
        if (!seen.has(ck[level])) {
          seen.add(ck[level]);
          const groupPrefix = [...prefix, ck[level]];
          const stKey = `_subtotal_${groupPrefix.join('\x01')}`;
          groups.push({
            headerName: fmtDimVal(ck[level], colDims[level]),
            children: [
              ...buildColGroup(keys, level + 1, groupPrefix),
              { headerName: 'Subtotal', field: stKey, sortable: true, type: 'numericColumn', valueFormatter } as ColDef,
            ],
          } as ColGroupDef);
        }
      }
      return groups;
    };

    const dimColDefs: ColDef[] = rowDims.map((d, i) => ({
      headerName: dimLabel(d),
      field: `_dim_${i}`,
      pinned: 'left' as const,
      sortable: true,
    }));

    const valueColDefs: (ColDef | ColGroupDef)[] = hasColDims
      ? buildColGroup(pivot.colKeys, 0, [])
      : [{
          headerName: config.aggregate === 'count' ? 'Count' : (config.yColumn ? dimLabel(config.yColumn) : 'Value'),
          field: '_val',
          sortable: true,
          type: 'numericColumn',
          valueFormatter,
        } as ColDef];

    const colDefs: (ColDef | ColGroupDef)[] = [
      ...dimColDefs,
      ...valueColDefs,
      { headerName: 'Total', field: '_total', sortable: true, type: 'numericColumn', valueFormatter } as ColDef,
    ];

    const pivotGridTheme = themeQuartz.withParams({
      cellHorizontalPaddingScale: 0.5,
      headerFontSize: 11,
      fontSize: 12,
      rowHeight: 24,
      headerHeight: 26,
    });

    return (
      <div style={{ height: '100%' }}>
        <AgGridReact
          theme={pivotGridTheme}
          modules={[AllCommunityModule]}
          rowData={rowData}
          columnDefs={colDefs}
          pinnedBottomRowData={[totalRow]}
          defaultColDef={{ resizable: true, suppressMovable: true }}
          suppressCellFocus
          suppressColumnVirtualisation
          onFirstDataRendered={e => e.api.autoSizeAllColumns()}
        />
      </div>
    );
  }

  if (data.length === 0) {
    return (
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%', color: 'var(--color-text-muted)', fontSize: 13 }}>
        No data
      </div>
    );
  }
  const xLabel = config.xLabel;
  const yLabel = config.yLabel;
  const margin = {
    top: 4,
    right: 16,
    bottom: xLabel ? 28 : 4,
    left: yLabel ? 16 : 0,
  };
  const xIsYearMonth = config.xColumn?.endsWith(':yearmonth') ?? false;
  const xTickFormatter = xIsYearMonth
    ? (v: string) => { const d = new Date(v); return isNaN(d.getTime()) ? v : d.toLocaleString('default', { month: 'short', year: 'numeric' }); }
    : undefined;
  const xAxisLabel = xLabel ? { value: xLabel, position: 'insideBottom' as const, offset: -8, fontSize: 11, fill: 'var(--color-text-muted)' } : undefined;
  const yAxisLabel = yLabel ? { value: yLabel, angle: -90, position: 'insideLeft' as const, offset: 8, fontSize: 11, fill: 'var(--color-text-muted)' } : undefined;
  const hasValueFormat = !!(config.valueFormat || config.yModifier);
  const yTickFormatter = hasValueFormat ? (v: number) => formatValue(v, config) : undefined;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const tooltipFormatter = hasValueFormat ? (v: any) => [formatValue(Number(v), config), undefined] : undefined;

  if (config.type === 'pie') {
    const pieData = data.map((d, i) => ({
      name: String(d.x ?? ''),
      value: Number(d[seriesKeys[0]] ?? 0),
      fill: CHART_COLORS[i % CHART_COLORS.length],
    }));
    return (
      <ResponsiveContainer width="100%" height="100%">
        <PieChart>
          <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius="70%">
            {pieData.map((entry, i) => <Cell key={i} fill={entry.fill} />)}
          </Pie>
          <Tooltip />
          <Legend />
        </PieChart>
      </ResponsiveContainer>
    );
  }

  if (config.type === 'scatter') {
    return (
      <ResponsiveContainer width="100%" height="100%">
        <ScatterChart margin={margin}>
          <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border)" />
          <XAxis dataKey="x" name={config.xColumn} tick={{ fontSize: 11 }} stroke="var(--color-border)" label={xAxisLabel} tickFormatter={xTickFormatter} />
          <YAxis dataKey={seriesKeys[0]} name={config.yColumn} tick={{ fontSize: 11 }} stroke="var(--color-border)" allowDecimals={false} label={yAxisLabel} tickFormatter={yTickFormatter} />
          <Tooltip cursor={{ strokeDasharray: '3 3' }} formatter={tooltipFormatter} />
          <Scatter data={data} fill={CHART_COLORS[0]} />
        </ScatterChart>
      </ResponsiveContainer>
    );
  }

  if (config.type === 'bar') {
    return (
      <ResponsiveContainer width="100%" height="100%">
        <BarChart data={data} margin={margin}>
          <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border)" />
          <XAxis dataKey="x" tick={{ fontSize: 11 }} stroke="var(--color-border)" label={xAxisLabel} tickFormatter={xTickFormatter} />
          <YAxis tick={{ fontSize: 11 }} stroke="var(--color-border)" allowDecimals={false} label={yAxisLabel} tickFormatter={yTickFormatter} />
          <Tooltip formatter={tooltipFormatter} />
          {seriesKeys.length > 1 && <Legend />}
          {seriesKeys.map((k, i) => (
            <Bar key={k} dataKey={k} fill={CHART_COLORS[i % CHART_COLORS.length]} radius={config.stacked ? undefined : [3, 3, 0, 0]} stackId={config.stacked ? 'stack' : undefined} />
          ))}
        </BarChart>
      </ResponsiveContainer>
    );
  }

  if (config.type === 'line') {
    return (
      <ResponsiveContainer width="100%" height="100%">
        <LineChart data={data} margin={margin}>
          <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border)" />
          <XAxis dataKey="x" tick={{ fontSize: 11 }} stroke="var(--color-border)" label={xAxisLabel} tickFormatter={xTickFormatter} />
          <YAxis tick={{ fontSize: 11 }} stroke="var(--color-border)" allowDecimals={false} label={yAxisLabel} tickFormatter={yTickFormatter} />
          <Tooltip formatter={tooltipFormatter} />
          {seriesKeys.length > 1 && <Legend />}
          {seriesKeys.map((k, i) => (
            <Line key={k} type="monotone" dataKey={k} stroke={CHART_COLORS[i % CHART_COLORS.length]} dot={false} strokeWidth={2} />
          ))}
        </LineChart>
      </ResponsiveContainer>
    );
  }

  // area / area-stacked (default)
  const areaStacked = config.type === 'area-stacked';
  return (
    <ResponsiveContainer width="100%" height="100%">
      <AreaChart data={data} margin={margin}>
        <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border)" />
        <XAxis dataKey="x" tick={{ fontSize: 11 }} stroke="var(--color-border)" label={xAxisLabel} tickFormatter={xTickFormatter} />
        <YAxis tick={{ fontSize: 11 }} stroke="var(--color-border)" allowDecimals={false} label={yAxisLabel} tickFormatter={yTickFormatter} />
        <Tooltip formatter={tooltipFormatter} />
        {seriesKeys.length > 1 && <Legend />}
        {seriesKeys.map((k, i) => (
          <Area
            key={k}
            type="monotone"
            dataKey={k}
            stroke={CHART_COLORS[i % CHART_COLORS.length]}
            fill={CHART_COLORS[i % CHART_COLORS.length] + '33'}
            strokeWidth={2}
            stackId={areaStacked ? 'stack' : undefined}
          />
        ))}
      </AreaChart>
    </ResponsiveContainer>
  );
};

// ── Config modal ─────────────────────────────────────────────────────────────

const FORMAT_HINT_ROWS: { tpl: string; desc: string }[] = [
  { tpl: '{value}',        desc: 'Raw number' },
  { tpl: '{value:.2f}',    desc: '2 decimal places  →  3.14' },
  { tpl: '{value:,.2f}',   desc: 'Thousands + 2 decimals  →  1,234.56' },
  { tpl: '{value:,}',      desc: 'Thousands separator  →  1,235' },
  { tpl: '{value:.2%}',    desc: 'Percentage  →  12.35%' },
  { tpl: '{value:.2s}',    desc: 'SI prefix  →  1.23k / 4.56M' },
  { tpl: '{value} km',     desc: 'Append a unit' },
  { tpl: '${value:,.2f}',  desc: 'Currency  →  $1,234.56' },
];

const FormatHint: React.FC = () => {
  const [open, setOpen] = React.useState(false);
  return (
    <div style={{ fontSize: 11, color: 'var(--color-text-muted)' }}>
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--color-text-muted)', fontSize: 11, textDecoration: 'underline dotted' }}
      >
        {open ? 'Hide examples' : 'Show format examples'}
      </button>
      {open && (
        <table style={{ marginTop: 6, borderCollapse: 'collapse', width: '100%' }}>
          <tbody>
            {FORMAT_HINT_ROWS.map(r => (
              <tr key={r.tpl}>
                <td style={{ fontFamily: 'monospace', paddingRight: 12, paddingBottom: 2, whiteSpace: 'nowrap' }}>{r.tpl}</td>
                <td style={{ paddingBottom: 2, color: 'var(--color-text-muted)' }}>{r.desc}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
};

const CHART_TYPE_OPTIONS = [
  { value: 'bar', label: 'Bar' },
  { value: 'line', label: 'Line' },
  { value: 'area', label: 'Area' },
  { value: 'area-stacked', label: 'Area (stacked)' },
  { value: 'pie', label: 'Pie' },
  { value: 'scatter', label: 'Scatter' },
  { value: 'table', label: 'Table' },
];

const AGGREGATE_OPTIONS = [
  { value: 'count', label: 'Count' },
  { value: 'sum', label: 'Sum' },
  { value: 'avg', label: 'Average' },
  { value: 'min', label: 'Min' },
  { value: 'max', label: 'Max' },
  { value: 'none', label: 'None (raw)' },
];

const DATE_FEATURES: { value: DateFeature; label: string }[] = [
  { value: 'year', label: 'Year' },
  { value: 'quarter', label: 'Quarter' },
  { value: 'yearmonth', label: 'Month' },
  { value: 'week', label: 'Week of year' },
  { value: 'dayofweek', label: 'Day of week' },
  { value: 'day', label: 'Day of month' },
  { value: 'hour', label: 'Hour' },
];

const ChartConfigModal: React.FC<{
  config: ChartConfig;
  isNew: boolean;
  tableIds: string[];
  getColumnPaths: (tableId: string) => { path: string; label: string; type?: string }[];
  getSchema: (tableId: string) => import('./types').TableSchema | undefined;
  updateSchema: (tableId: string, schema: import('./types').TableSchema) => Promise<void>;
  onSave: (config: ChartConfig) => void;
  onDelete?: () => void;
  onClose: () => void;
}> = ({ config, isNew, tableIds, getColumnPaths, getSchema, updateSchema, onSave, onDelete, onClose }) => {
  const [draft, setDraft] = useState<ChartConfig>(config);
  const allPaths = getColumnPaths(draft.table);
  // Only leaf paths (no further ref children) or non-ref columns are valid Y columns
  const numericPathSet = new Set(
    allPaths
      .filter(p => !p.type || (p.type !== 'reference' && p.type !== 'image' && p.type !== 'bool'))
      .map(p => p.path)
  );
  const isTableType = draft.type === 'table';
  const hasGroupBy = !isTableType && (draft.type === 'bar' || draft.type === 'line' || draft.type === 'area' || draft.type === 'area-stacked');
  const needsYCol = draft.aggregate !== 'count' && draft.aggregate !== 'none';
  const canSave = isTableType
    ? (draft.tableRows?.length ?? 0) > 0 && (!needsYCol || !!draft.yColumn)
    : !!draft.xColumn && (!needsYCol || !!draft.yColumn);

  type ColOption = { value: string; label: string };
  type ColGroup = { label: string; options: ColOption[] };

  // ── colOptions for tableRows/tableColumns: date cols have date-feature sub-options ──
  const colGroupsMap = new Map<string, ColGroup>();
  const directGroup: ColGroup = { label: 'Columns', options: [] };
  for (const p of allPaths) {
    const isDotPath = p.path.includes('.');
    const groupKey = isDotPath ? p.path.split('.')[0] : null;
    if (isDotPath && groupKey) {
      if (!colGroupsMap.has(groupKey)) colGroupsMap.set(groupKey, { label: groupKey, options: [] });
      const shortLabel = p.label.includes(' → ') ? p.label.split(' → ').slice(1).join(' → ') : p.label;
      colGroupsMap.get(groupKey)!.options.push({ value: p.path, label: shortLabel });
    } else if (p.type === 'date' || p.type === 'datetime') {
      const feats = p.type === 'date' ? DATE_FEATURES.filter(f => f.value !== 'hour') : DATE_FEATURES;
      const dateGroup: ColGroup = { label: p.label, options: [{ value: p.path, label: 'Raw value' }, ...feats.map(f => ({ value: `${p.path}:${f.value}`, label: f.label }))] };
      colGroupsMap.set(p.path, dateGroup);
    } else {
      directGroup.options.push({ value: p.path, label: p.label });
    }
  }
  const colOptions: ColGroup[] = [];
  if (directGroup.options.length > 0) colOptions.push(directGroup);
  for (const g of colGroupsMap.values()) colOptions.push(g);
  const colOptionsFlat: ColOption[] = colOptions.flatMap(g => g.options);

  // ── colOptionsFlat for X/Y/groupBy: date cols appear as a single selectable option ──
  const directGroupXYG: ColGroup = { label: 'Columns', options: [] };
  const refGroupsXYG = new Map<string, ColGroup>();
  for (const p of allPaths) {
    const isDot = p.path.includes('.');
    const rootKey = isDot ? p.path.split('.')[0] : null;
    if (isDot && rootKey) {
      if (!refGroupsXYG.has(rootKey)) refGroupsXYG.set(rootKey, { label: rootKey, options: [] });
      const short = p.label.includes(' → ') ? p.label.split(' → ').slice(1).join(' → ') : p.label;
      refGroupsXYG.get(rootKey)!.options.push({ value: p.path, label: short });
    } else {
      directGroupXYG.options.push({ value: p.path, label: p.label });
    }
  }
  const colOptionsXYG: ColGroup[] = [];
  if (directGroupXYG.options.length > 0) colOptionsXYG.push(directGroupXYG);
  for (const g of refGroupsXYG.values()) colOptionsXYG.push(g);
  const colOptionsFlatXYG: ColOption[] = colOptionsXYG.flatMap(g => g.options);
  const yColOptionsXYG: ColGroup[] = draft.aggregate === 'none'
    ? colOptionsXYG
    : colOptionsXYG.map(g => ({ ...g, options: g.options.filter(o => numericPathSet.has(o.value)) })).filter(g => g.options.length > 0);

  // ── Filtered grouped options for Y column in table mode ──
  const yColOptions: ColGroup[] = draft.aggregate === 'none'
    ? colOptions
    : colOptions.map(g => ({ ...g, options: g.options.filter(o => numericPathSet.has(o.value)) })).filter(g => g.options.length > 0);

  // Parse date feature out of xColumn / groupBy
  const xColPath = stripFeature(draft.xColumn);
  const xDateFeature = getFeature(draft.xColumn);
  const xColType = allPaths.find(p => p.path === xColPath)?.type;
  const xIsDateCol = xColType === 'date' || xColType === 'datetime';

  const gbColPath = draft.groupBy ? stripFeature(draft.groupBy) : '';
  const gbDateFeature = draft.groupBy ? getFeature(draft.groupBy) : undefined;
  const gbColType = allPaths.find(p => p.path === gbColPath)?.type;
  const gbIsDateCol = gbColType === 'date' || gbColType === 'datetime';

  const set = <K extends keyof ChartConfig>(key: K, val: ChartConfig[K]) =>
    setDraft(d => ({ ...d, [key]: val }));

  const modSectionStyle: React.CSSProperties = { display: 'flex', flexDirection: 'column', gap: 6, padding: '8px 10px', background: 'var(--color-surface-raised, rgba(0,0,0,0.04))', borderRadius: 6, border: '1px solid var(--color-border)' };
  const modLabelStyle: React.CSSProperties = { fontSize: 11, fontWeight: 700, color: 'var(--color-text-muted)', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 2 };

  return (
    <div
      style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 3000 }}
      onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div style={{ background: 'var(--color-surface)', border: '1px solid var(--color-border)', borderRadius: 12, width: 'min(480px, 94vw)', maxHeight: '90vh', display: 'flex', flexDirection: 'column', boxShadow: '0 12px 40px rgba(0,0,0,0.22)', color: 'var(--color-text)', overflow: 'hidden' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '14px 20px', borderBottom: '1px solid var(--color-border)' }}>
          <span style={{ fontWeight: 600, fontSize: 15 }}>{isNew ? 'Add Chart' : 'Edit Chart'}</span>
          <button onClick={onClose} className="app-dialog-close" aria-label="Close">×</button>
        </div>
        <div style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: 12, overflowY: 'auto' }}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            <label className="app-dialog-label" style={{ marginBottom: 0 }}>Title</label>
            <input
              className="app-dialog-input"
              style={{ marginBottom: 0 }}
              value={draft.title}
              onChange={e => set('title', e.target.value)}
              placeholder="Chart title"
            />
          </div>
          <div style={{ display: 'flex', gap: 12 }}>
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
              <label className="app-dialog-label" style={{ marginBottom: 0 }}>Type</label>
              <Select
                styles={dialogSelectStyles}
                isSearchable={false}
                value={CHART_TYPE_OPTIONS.find(o => o.value === draft.type) ?? null}
                options={CHART_TYPE_OPTIONS}
                onChange={opt => set('type', (opt?.value ?? 'bar') as ChartType)}
                menuPortalTarget={document.body}
                menuPlacement="auto"
              />
            </div>
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
              <label className="app-dialog-label" style={{ marginBottom: 0 }}>Table</label>
              <Select
                styles={dialogSelectStyles}
                isSearchable={false}
                value={draft.table ? { value: draft.table, label: draft.table } : null}
                options={tableIds.map(id => ({ value: id, label: id }))}
                onChange={opt => { set('table', opt?.value ?? ''); set('xColumn', ''); set('yColumn', ''); }}
                menuPortalTarget={document.body}
                menuPlacement="auto"
              />
            </div>
          </div>
          {isTableType ? (
            <>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                <label className="app-dialog-label" style={{ marginBottom: 0 }}>Rows</label>
                <Select
                  styles={dialogSelectStyles}
                  isMulti
                  value={(draft.tableRows ?? []).map(v => colOptionsFlat.find(o => o.value === v) ?? { value: v, label: v })}
                  options={colOptions}
                  onChange={opts => set('tableRows', opts ? opts.map(o => o.value) : [])}
                  placeholder="— select row dimensions —"
                  menuPortalTarget={document.body}
                  menuPlacement="auto"
                />
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                <label className="app-dialog-label" style={{ marginBottom: 0 }}>
                  Columns <span style={{ fontWeight: 400, color: 'var(--color-text-muted)' }}>(optional)</span>
                </label>
                <Select
                  styles={dialogSelectStyles}
                  isMulti
                  value={(draft.tableColumns ?? []).map(v => colOptionsFlat.find(o => o.value === v) ?? { value: v, label: v })}
                  options={colOptions}
                  onChange={opts => set('tableColumns', opts ? opts.map(o => o.value) : [])}
                  placeholder="— none —"
                  menuPortalTarget={document.body}
                  menuPlacement="auto"
                />
              </div>
              <div style={{ display: 'flex', gap: 12 }}>
                <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <label className="app-dialog-label" style={{ marginBottom: 0 }}>Aggregate</label>
                  <Select
                    styles={dialogSelectStyles}
                    isSearchable={false}
                    value={AGGREGATE_OPTIONS.find(o => o.value === draft.aggregate) ?? null}
                    options={AGGREGATE_OPTIONS}
                    onChange={opt => set('aggregate', (opt?.value ?? 'sum') as AggregateFunc)}
                    menuPortalTarget={document.body}
                    menuPlacement="auto"
                  />
                </div>
                {needsYCol && (
                  <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <label className="app-dialog-label" style={{ marginBottom: 0 }}>Value column</label>
                    <Select
                      styles={dialogSelectStyles}
                      value={colOptionsFlat.find(o => o.value === draft.yColumn) ?? null}
                      options={yColOptions}
                      onChange={opt => set('yColumn', opt?.value ?? '')}
                      placeholder="— select —"
                      isClearable
                      menuPortalTarget={document.body}
                      menuPlacement="auto"
                    />
                  </div>
                )}
              </div>
              {needsYCol && draft.yColumn && (
                <div style={modSectionStyle}>
                  <div style={modLabelStyle}>Value options</div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                      <label className="app-dialog-label" style={{ marginBottom: 0, fontWeight: 400 }}>Display format <span style={{ color: 'var(--color-text-muted)', fontWeight: 400 }}>(optional)</span></label>
                      <input type="text" className="app-dialog-input" style={{ marginBottom: 0, fontFamily: 'monospace', fontSize: 12 }} placeholder="{{value}}" value={draft.valueFormat ?? ''} onChange={e => set('valueFormat', e.target.value || undefined)} />
                    </div>
                    <FormatHint />
                  </div>
                </div>
              )}
            </>
          ) : (
            <>
              <div style={{ display: 'flex', gap: 12 }}>
                <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <label className="app-dialog-label" style={{ marginBottom: 0 }}>X column</label>
                  <Select
                    styles={dialogSelectStyles}
                    value={colOptionsFlatXYG.find(o => o.value === xColPath) ?? null}
                    options={colOptionsXYG}
                    onChange={opt => {
                      const newPath = opt?.value ?? '';
                      const newType = allPaths.find(p => p.path === newPath)?.type;
                      if (newType === 'date' || newType === 'datetime') {
                        set('xColumn', xDateFeature ? `${newPath}:${xDateFeature}` : newPath);
                      } else {
                        set('xColumn', newPath);
                      }
                    }}
                    placeholder="— select —"
                    isClearable
                    menuPortalTarget={document.body}
                    menuPlacement="auto"
                  />
                  {xIsDateCol && (
                    <Select
                      styles={dialogSelectStyles}
                      isSearchable={false}
                      placeholder="Date feature (raw)"
                      isClearable
                      value={DATE_FEATURES.find(f => f.value === xDateFeature) ?? null}
                      options={xColType === 'date' ? DATE_FEATURES.filter(f => f.value !== 'hour') : DATE_FEATURES}
                      onChange={opt => set('xColumn', opt?.value ? `${xColPath}:${opt.value}` : xColPath)}
                      menuPortalTarget={document.body}
                      menuPlacement="auto"
                    />
                  )}
                </div>
                <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <label className="app-dialog-label" style={{ marginBottom: 0 }}>Aggregate</label>
                  <Select
                    styles={dialogSelectStyles}
                    isSearchable={false}
                    value={AGGREGATE_OPTIONS.find(o => o.value === draft.aggregate) ?? null}
                    options={AGGREGATE_OPTIONS}
                    onChange={opt => set('aggregate', (opt?.value ?? 'count') as AggregateFunc)}
                    menuPortalTarget={document.body}
                    menuPlacement="auto"
                  />
                </div>
              </div>
              {needsYCol && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <label className="app-dialog-label" style={{ marginBottom: 0 }}>Y column</label>
                  <Select
                    styles={dialogSelectStyles}
                    value={colOptionsFlatXYG.find(o => o.value === draft.yColumn) ?? null}
                    options={yColOptionsXYG}
                    onChange={opt => set('yColumn', opt?.value ?? '')}
                    placeholder="— select —"
                    isClearable
                    menuPortalTarget={document.body}
                    menuPlacement="auto"
                  />
                </div>
              )}
              {needsYCol && draft.yColumn && (
                <div style={modSectionStyle}>
                  <div style={modLabelStyle}>Value options</div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                      <label className="app-dialog-label" style={{ marginBottom: 0, fontWeight: 400 }}>Display format <span style={{ color: 'var(--color-text-muted)', fontWeight: 400 }}>(optional)</span></label>
                      <input type="text" className="app-dialog-input" style={{ marginBottom: 0, fontFamily: 'monospace', fontSize: 12 }} placeholder="{{value}}" value={draft.valueFormat ?? ''} onChange={e => set('valueFormat', e.target.value || undefined)} />
                    </div>
                    <FormatHint />
                  </div>
                </div>
              )}
              {draft.type === 'bar' && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <input
                    type="checkbox"
                    id="chart-stacked-cb"
                    checked={!!draft.stacked}
                    onChange={e => set('stacked', e.target.checked || undefined)}
                  />
                  <label htmlFor="chart-stacked-cb" className="app-dialog-label" style={{ marginBottom: 0 }}>Stacked bars</label>
                </div>
              )}
              {hasGroupBy && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <label className="app-dialog-label" style={{ marginBottom: 0 }}>
                    Group by <span style={{ fontWeight: 400, color: 'var(--color-text-muted)' }}>(optional)</span>
                  </label>
                  <Select
                    styles={dialogSelectStyles}
                    value={gbColPath ? colOptionsFlatXYG.find(o => o.value === gbColPath) ?? null : null}
                    options={colOptionsXYG}
                    onChange={opt => {
                      const newPath = opt?.value;
                      if (!newPath) { set('groupBy', undefined); return; }
                      const newType = allPaths.find(p => p.path === newPath)?.type;
                      if (newType === 'date' || newType === 'datetime') {
                        set('groupBy', gbDateFeature ? `${newPath}:${gbDateFeature}` : newPath);
                      } else {
                        set('groupBy', newPath);
                      }
                    }}
                    placeholder="— none —"
                    isClearable
                    menuPortalTarget={document.body}
                    menuPlacement="auto"
                  />
                  {gbIsDateCol && gbColPath && (
                    <Select
                      styles={dialogSelectStyles}
                      isSearchable={false}
                      placeholder="Date feature (raw)"
                      isClearable
                      value={DATE_FEATURES.find(f => f.value === gbDateFeature) ?? null}
                      options={gbColType === 'date' ? DATE_FEATURES.filter(f => f.value !== 'hour') : DATE_FEATURES}
                      onChange={opt => set('groupBy', opt?.value ? `${gbColPath}:${opt.value}` : gbColPath)}
                      menuPortalTarget={document.body}
                      menuPlacement="auto"
                    />
                  )}
                </div>
              )}
            </>
          )}
          {draft.table && (() => {
            const schema = getSchema(draft.table);
            if (!schema) return null;
            const calcCols = schema.calculatedColumns ?? [];
            return (
              <div style={modSectionStyle}>
                <div style={modLabelStyle}>Calculated columns</div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                  {calcCols.map((calc, idx) => (
                    <div key={idx} style={{ display: 'flex', gap: 6, alignItems: 'flex-start' }}>
                      <div style={{ flex: '0 0 110px', display: 'flex', flexDirection: 'column', gap: 2 }}>
                        <input
                          className="app-dialog-input"
                          style={{ marginBottom: 0, fontSize: 12 }}
                          placeholder="name"
                          value={calc.name}
                          onChange={e => {
                            const updated = calcCols.map((c, i) => i === idx ? { ...c, name: e.target.value } : c);
                            void updateSchema(draft.table, { ...schema, calculatedColumns: updated });
                          }}
                        />
                        <input
                          className="app-dialog-input"
                          style={{ marginBottom: 0, fontSize: 11 }}
                          placeholder="display name"
                          value={calc.displayName ?? ''}
                          onChange={e => {
                            const updated = calcCols.map((c, i) => i === idx ? { ...c, displayName: e.target.value || undefined } : c);
                            void updateSchema(draft.table, { ...schema, calculatedColumns: updated });
                          }}
                        />
                      </div>
                      <input
                        className="app-dialog-input"
                        style={{ marginBottom: 0, fontFamily: 'monospace', fontSize: 12, flex: 1 }}
                        placeholder="expression (e.g. distance / 1000)"
                        value={calc.expression}
                        onChange={e => {
                          const updated = calcCols.map((c, i) => i === idx ? { ...c, expression: e.target.value } : c);
                          void updateSchema(draft.table, { ...schema, calculatedColumns: updated });
                        }}
                      />
                      <button
                        className="app-dialog-close"
                        style={{ flexShrink: 0 }}
                        onClick={() => {
                          const updated = calcCols.filter((_, i) => i !== idx);
                          void updateSchema(draft.table, { ...schema, calculatedColumns: updated });
                        }}
                        title="Remove"
                      >×</button>
                    </div>
                  ))}
                  <button
                    className="btn-secondary btn-sm"
                    style={{ alignSelf: 'flex-start' }}
                    onClick={() => {
                      const updated = [...calcCols, { name: '', expression: '' }];
                      void updateSchema(draft.table, { ...schema, calculatedColumns: updated });
                    }}
                  >+ Add column</button>
                </div>
              </div>
            );
          })()}
          {draft.type !== 'pie' && draft.type !== 'table' && (
            <div style={{ display: 'flex', gap: 12 }}>
              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
                <label className="app-dialog-label" style={{ marginBottom: 0 }}>
                  X axis label <span style={{ fontWeight: 400, color: 'var(--color-text-muted)' }}>(optional)</span>
                </label>
                <input
                  className="app-dialog-input"
                  style={{ marginBottom: 0 }}
                  value={draft.xLabel ?? ''}
                  onChange={e => set('xLabel', e.target.value || undefined)}
                  placeholder="X axis label"
                />
              </div>
              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
                <label className="app-dialog-label" style={{ marginBottom: 0 }}>
                  Y axis label <span style={{ fontWeight: 400, color: 'var(--color-text-muted)' }}>(optional)</span>
                </label>
                <input
                  className="app-dialog-input"
                  style={{ marginBottom: 0 }}
                  value={draft.yLabel ?? ''}
                  onChange={e => set('yLabel', e.target.value || undefined)}
                  placeholder="Y axis label"
                />
              </div>
            </div>
          )}
        </div>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, padding: '12px 20px', borderTop: '1px solid var(--color-border)' }}>
          {!isNew && onDelete && (
            <button className="btn-danger" style={{ marginRight: 'auto' }} onClick={onDelete}>Delete chart</button>
          )}
          <button className="btn-secondary" onClick={onClose}>Cancel</button>
          <button className="btn-primary" onClick={() => onSave(draft)} disabled={!canSave}>
            {isNew ? 'Add Chart' : 'Save'}
          </button>
        </div>
      </div>
    </div>
  );
};

// ── Main page ─────────────────────────────────────────────────────────────────

export const ChartSheetPage: React.FC<{ state: UseAppStateReturn }> = ({ state }) => {
  const { bookId, chartId } = useParams<{ bookId?: string; chartId: string }>();
  const navigate = useNavigate();
  const confirm = useConfirm();

  const chartSheet = chartId ? state.getChartSheet(chartId) : undefined;

  const [charts, setCharts] = useState<ChartConfig[]>(() => chartSheet?.charts ?? []);
  const [layout, setLayout] = useState<ChartLayoutItem[]>(() => chartSheet?.layout ?? []);
  const [editingChart, setEditingChart] = useState<ChartConfig | null>(null);
  const [isNewChart, setIsNewChart] = useState(false);

  const canEdit = state.activeBookRole === 'owner' || state.activeBookRole === 'editor';
  const [searchParams] = useSearchParams();
  const editLayout = searchParams.get('editLayout') === '1';

  // Sync local state from app state when chartSheet first loads (async) or when navigating to a different chart
  useEffect(() => {
    if (chartSheet) {
      setCharts(chartSheet.charts);
      setLayout(chartSheet.layout);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chartId, !!chartSheet]);

  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const scheduleSave = useCallback((nextCharts: ChartConfig[], nextLayout: ChartLayoutItem[]) => {
    if (!chartId) return;
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      void state.updateChartSheet(chartId, nextCharts, nextLayout);
    }, 800);
  }, [chartId, state]);

  const handleLayoutChange = useCallback((newLayout: readonly ChartLayoutItem[]) => {
    const nextLayout: ChartLayoutItem[] = newLayout.map(({ i, x, y, w, h }) => ({ i, x, y, w, h }));
    setLayout(nextLayout);
    setCharts(prev => { scheduleSave(prev, nextLayout); return prev; });
  }, [scheduleSave]);

  const handleAddChart = () => {
    const firstTable = state.tableIds[0] ?? '';
    const schema = state.getSchema(firstTable);
    const cols = schema?.columns ?? [];
    setEditingChart({
      id: crypto.randomUUID(),
      title: 'New Chart',
      type: 'bar',
      table: firstTable,
      xColumn: cols[0]?.name ?? '',
      yColumn: cols.find(c => c.type === 'integer' || c.type === 'decimal')?.name ?? '',
      aggregate: 'sum',
      tableRows: [],
      tableColumns: [],
    });
    setIsNewChart(true);
  };

  const handleSaveChart = useCallback((config: ChartConfig) => {
    setEditingChart(null);
    setCharts(prevCharts => {
      setLayout(prevLayout => {
        const nextCharts = isNewChart ? [...prevCharts, config] : prevCharts.map(c => c.id === config.id ? config : c);
        const nextLayout = isNewChart
          ? [...prevLayout, { i: config.id, x: (prevLayout.length * 6) % 12, y: Infinity, w: 6, h: 8 }]
          : prevLayout;
        scheduleSave(nextCharts, nextLayout);
        return nextLayout;
      });
      return isNewChart ? [...prevCharts, config] : prevCharts.map(c => c.id === config.id ? config : c);
    });
  }, [isNewChart, scheduleSave]);

  const handleDeleteChart = useCallback((id: string) => {
    setCharts(prevCharts => {
      setLayout(prevLayout => {
        const nextCharts = prevCharts.filter(c => c.id !== id);
        const nextLayout = prevLayout.filter(l => l.i !== id);
        scheduleSave(nextCharts, nextLayout);
        return nextLayout;
      });
      return prevCharts.filter(c => c.id !== id);
    });
  }, [scheduleSave]);

  const handleDeleteSheet = useCallback(async () => {
    if (!chartId) return;
    const confirmed = await confirm(`Delete chart sheet "${chartId}"?`, 'Delete');
    if (!confirmed) return;
    await state.deleteChartSheet(chartId);
    const dest = state.tableIds[0]
      ? `${bookPrefix(bookId)}/table/${encodeURIComponent(state.tableIds[0])}`
      : bookPrefix(bookId);
    navigate(dest, { replace: true });
  }, [bookId, chartId, confirm, navigate, state]);

  const getColumnPathsForTable = useCallback((tableId: string) =>
    state.getColumnPaths(tableId).map(p => {
      // Attach type for the leaf column so date features can be offered
      const leafCol = p.path.split('.').reduce<string | null>((tableName, col) => {
        if (!tableName) return null;
        const schema = state.getSchema(tableName);
        const colDef = schema?.columns.find(c => c.name === col);
        return colDef?.type === 'reference' ? (colDef.refTable ?? null) : null;
      }, tableId);
      // Determine leaf column type
      const parts = p.path.split('.');
      let t = tableId;
      let leafType: string | undefined;
      for (let i = 0; i < parts.length; i++) {
        const s = state.getSchema(t);
        const cd = s?.columns.find(c => c.name === parts[i]);
        if (!cd) break;
        if (i === parts.length - 1) { leafType = cd.type; break; }
        if (cd.type === 'reference' && cd.refTable) t = cd.refTable; else break;
      }
      void leafCol;
      return { path: p.path, label: p.label, type: leafType };
    }),
  [state]); // eslint-disable-line react-hooks/exhaustive-deps

  if (!chartId || !chartSheet) {
    return (
      <div className="app-body">
        <div className="main-content">
          <div className="empty-state-main"><h2>Chart sheet not found</h2></div>
        </div>
      </div>
    );
  }

  return (
    <div className="app-body" style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      {editingChart && (
        <ChartConfigModal
          config={editingChart}
          isNew={isNewChart}
          tableIds={state.tableIds}
          getColumnPaths={getColumnPathsForTable}
          getSchema={state.getSchema}
          updateSchema={state.updateSchema}
          onSave={handleSaveChart}
          onDelete={!isNewChart ? () => { handleDeleteChart(editingChart.id); setEditingChart(null); } : undefined}
          onClose={() => setEditingChart(null)}
        />
      )}
      <div style={{ flex: 1, overflowY: 'auto', padding: 16 }}>
        {charts.length === 0 ? (
          <div className="empty-state-main" style={{ minHeight: 320 }}>
            <h2>No charts yet</h2>
            {canEdit && state.tableIds.length > 0 && (
              <button className="btn-primary" onClick={handleAddChart}>Add your first chart</button>
            )}
            {state.tableIds.length === 0 && (
              <p style={{ color: 'var(--color-text-muted)' }}>Create a table first to add charts.</p>
            )}
          </div>
        ) : (
          <RGL
            layout={layout}
            cols={12}
            rowHeight={40}
            margin={[12, 12]}
            onLayoutChange={handleLayoutChange}
            draggableHandle=".chart-drag-handle"
            isDraggable={canEdit && editLayout}
            isResizable={canEdit && editLayout}
          >
            {charts.map(chart => {
              const rows = state.getRows(chart.table);
              const { data, seriesKeys } = chart.type === 'table'
                ? { data: [], seriesKeys: [] }
                : aggregateData(rows, chart.xColumn, chart.yColumn, chart.aggregate, chart.groupBy, state.resolveColumnPath, chart.table);
              return (
                <div
                  key={chart.id}
                  style={{
                    background: 'var(--color-surface)',
                    border: '1px solid var(--color-border)',
                    borderRadius: 8,
                    display: 'flex',
                    flexDirection: 'column',
                    overflow: 'hidden',
                  }}
                >
                  <div
                    className="chart-drag-handle"
                    style={{
                      padding: '8px 12px',
                      borderBottom: '1px solid var(--color-border)',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      cursor: editLayout ? 'grab' : 'default',
                      flexShrink: 0,
                      userSelect: 'none',
                    }}
                  >
                    <span style={{ fontWeight: 600, fontSize: 13, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {chart.title}
                    </span>
                    {canEdit && (
                      <div style={{ display: 'flex', gap: 4, flexShrink: 0 }} onMouseDown={e => e.stopPropagation()} onTouchStart={e => e.stopPropagation()}>
                        <button
                          onClick={() => { setEditingChart(chart); setIsNewChart(false); }}
                          style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--color-text-muted)', padding: '2px 6px', fontSize: 12, borderRadius: 4 }}
                        >Edit</button>
                      </div>
                    )}
                  </div>
                  <div style={{ flex: 1, minHeight: 0, padding: '8px 4px 4px' }}>
                    <ChartRenderer
                      config={chart}
                      data={data}
                      seriesKeys={seriesKeys}
                      rows={rows}
                      resolveColumnPath={state.resolveColumnPath}
                      getColumnPaths={getColumnPathsForTable}
                      onSortChange={sort => {
                        setCharts(prev => {
                          const next = prev.map(c => c.id === chart.id ? { ...c, tableSort: sort } : c);
                          scheduleSave(next, layout);
                          return next;
                        });
                      }}
                    />
                  </div>
                </div>
              );
            })}
          </RGL>
        )}
      </div>
      {canEdit && (
        <div style={{ padding: '8px 16px', borderTop: '1px solid var(--color-border)', display: 'flex', gap: 8, alignItems: 'center', flexShrink: 0 }}>
          {state.tableIds.length > 0 && (
            <button className="btn-primary btn-sm" onClick={handleAddChart}>+ Add chart</button>
          )}
          <button
            className="btn-secondary btn-sm"
            style={{ color: 'var(--color-danger)', marginLeft: 'auto' }}
            onClick={() => void handleDeleteSheet()}
          >Delete sheet</button>
        </div>
      )}
    </div>
  );
};
