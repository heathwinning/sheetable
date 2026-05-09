import React, { useState, useCallback, useEffect, useRef } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
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
import type { UseAppStateReturn } from './useAppState';
import type { ChartConfig, ChartLayoutItem, ChartType, AggregateFunc, Row, DateFeature } from './types';
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
    const validFeatures: DateFeature[] = ['year', 'quarter', 'month', 'monthnum', 'week', 'dayofweek', 'day', 'hour'];
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

// ── Chart renderer ───────────────────────────────────────────────────────────

const ChartRenderer: React.FC<{
  config: ChartConfig;
  data: Record<string, unknown>[];
  seriesKeys: string[];
}> = ({ config, data, seriesKeys }) => {
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
  const xAxisLabel = xLabel ? { value: xLabel, position: 'insideBottom' as const, offset: -8, fontSize: 11, fill: 'var(--color-text-muted)' } : undefined;
  const yAxisLabel = yLabel ? { value: yLabel, angle: -90, position: 'insideLeft' as const, offset: 8, fontSize: 11, fill: 'var(--color-text-muted)' } : undefined;

  if (config.type === 'table') {
    const valueCol = seriesKeys[0] ?? 'value';
    const isGrouped = seriesKeys.length > 1;
    return (
      <div style={{ overflow: 'auto', height: '100%', fontSize: 12 }}>
        <table style={{ borderCollapse: 'collapse', width: '100%', minWidth: 160 }}>
          <thead>
            <tr>
              <th style={{ padding: '4px 10px', borderBottom: '2px solid var(--color-border)', textAlign: 'left', fontWeight: 600, color: 'var(--color-text-muted)', whiteSpace: 'nowrap' }}>
                {config.xColumn.includes(':') ? config.xColumn.replace(':', ' › ') : config.xColumn}
              </th>
              {isGrouped
                ? seriesKeys.map(k => (
                  <th key={k} style={{ padding: '4px 10px', borderBottom: '2px solid var(--color-border)', textAlign: 'right', fontWeight: 600, color: 'var(--color-text-muted)', whiteSpace: 'nowrap' }}>{k}</th>
                ))
                : <th style={{ padding: '4px 10px', borderBottom: '2px solid var(--color-border)', textAlign: 'right', fontWeight: 600, color: 'var(--color-text-muted)', whiteSpace: 'nowrap' }}>
                    {config.aggregate === 'count' ? 'Count' : `${config.aggregate}(${valueCol})`}
                  </th>
              }
            </tr>
          </thead>
          <tbody>
            {data.map((row, i) => (
              <tr key={i} style={{ background: i % 2 === 0 ? 'transparent' : 'var(--color-surface-raised, rgba(0,0,0,0.03))' }}>
                <td style={{ padding: '3px 10px', borderBottom: '1px solid var(--color-border)', whiteSpace: 'nowrap' }}>{String(row.x ?? '')}</td>
                {isGrouped
                  ? seriesKeys.map(k => (
                    <td key={k} style={{ padding: '3px 10px', borderBottom: '1px solid var(--color-border)', textAlign: 'right', tabularNums: true } as React.CSSProperties}>{String(row[k] ?? '')}</td>
                  ))
                  : <td style={{ padding: '3px 10px', borderBottom: '1px solid var(--color-border)', textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{String(row[valueCol] ?? '')}</td>
                }
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

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
          <XAxis dataKey="x" name={config.xColumn} tick={{ fontSize: 11 }} stroke="var(--color-border)" label={xAxisLabel} />
          <YAxis dataKey={seriesKeys[0]} name={config.yColumn} tick={{ fontSize: 11 }} stroke="var(--color-border)" allowDecimals={false} label={yAxisLabel} />
          <Tooltip cursor={{ strokeDasharray: '3 3' }} />
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
          <XAxis dataKey="x" tick={{ fontSize: 11 }} stroke="var(--color-border)" label={xAxisLabel} />
          <YAxis tick={{ fontSize: 11 }} stroke="var(--color-border)" allowDecimals={false} label={yAxisLabel} />
          <Tooltip />
          {seriesKeys.length > 1 && <Legend />}
          {seriesKeys.map((k, i) => (
            <Bar key={k} dataKey={k} fill={CHART_COLORS[i % CHART_COLORS.length]} radius={[3, 3, 0, 0]} />
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
          <XAxis dataKey="x" tick={{ fontSize: 11 }} stroke="var(--color-border)" label={xAxisLabel} />
          <YAxis tick={{ fontSize: 11 }} stroke="var(--color-border)" allowDecimals={false} label={yAxisLabel} />
          <Tooltip />
          {seriesKeys.length > 1 && <Legend />}
          {seriesKeys.map((k, i) => (
            <Line key={k} type="monotone" dataKey={k} stroke={CHART_COLORS[i % CHART_COLORS.length]} dot={false} strokeWidth={2} />
          ))}
        </LineChart>
      </ResponsiveContainer>
    );
  }

  // area (default)
  return (
    <ResponsiveContainer width="100%" height="100%">
      <AreaChart data={data} margin={margin}>
        <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border)" />
        <XAxis dataKey="x" tick={{ fontSize: 11 }} stroke="var(--color-border)" label={xAxisLabel} />
        <YAxis tick={{ fontSize: 11 }} stroke="var(--color-border)" allowDecimals={false} label={yAxisLabel} />
        <Tooltip />
        {seriesKeys.length > 1 && <Legend />}
        {seriesKeys.map((k, i) => (
          <Area
            key={k}
            type="monotone"
            dataKey={k}
            stroke={CHART_COLORS[i % CHART_COLORS.length]}
            fill={CHART_COLORS[i % CHART_COLORS.length] + '33'}
            strokeWidth={2}
          />
        ))}
      </AreaChart>
    </ResponsiveContainer>
  );
};

// ── Config modal ─────────────────────────────────────────────────────────────

const CHART_TYPE_OPTIONS = [
  { value: 'bar', label: 'Bar' },
  { value: 'line', label: 'Line' },
  { value: 'area', label: 'Area' },
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

const ChartConfigModal: React.FC<{
  config: ChartConfig;
  isNew: boolean;
  tableIds: string[];
  getColumnPaths: (tableId: string) => { path: string; label: string; type?: string }[];
  onSave: (config: ChartConfig) => void;
  onClose: () => void;
}> = ({ config, isNew, tableIds, getColumnPaths, onSave, onClose }) => {
  const [draft, setDraft] = useState<ChartConfig>(config);
  const allPaths = getColumnPaths(draft.table);
  // Only leaf paths (no further ref children) or non-ref columns are valid Y columns
  const numericPathSet = new Set(
    allPaths
      .filter(p => !p.type || (p.type !== 'reference' && p.type !== 'image' && p.type !== 'bool'))
      .map(p => p.path)
  );
  const hasGroupBy = draft.type === 'bar' || draft.type === 'line' || draft.type === 'area' || draft.type === 'table';
  const needsYCol = draft.aggregate !== 'count' && draft.type !== 'table';
  const canSave = !!draft.xColumn && (!needsYCol || !!draft.yColumn);

  // Build expanded column options — date/datetime leaf columns get date-feature sub-options
  const DATE_FEATURES: { value: DateFeature; label: string }[] = [
    { value: 'year', label: 'Year' },
    { value: 'quarter', label: 'Quarter' },
    { value: 'month', label: 'Month name' },
    { value: 'monthnum', label: 'Month #' },
    { value: 'week', label: 'Week of year' },
    { value: 'dayofweek', label: 'Day of week' },
    { value: 'day', label: 'Day of month' },
    { value: 'hour', label: 'Hour' },
  ];
  type ColOption = { value: string; label: string };
  type ColGroup = { label: string; options: ColOption[] };

  // Build grouped column options:
  //   - direct columns (no dot in path, no colon feature) stay in a "Columns" group
  //   - date columns and their feature sub-options form their own group
  //   - ref columns and their dot-path children form a group per ref column
  const colGroupsMap = new Map<string, ColGroup>();
  const directGroup: ColGroup = { label: 'Columns', options: [] };

  for (const p of allPaths) {
    const isDotPath = p.path.includes('.');
    const groupKey = isDotPath ? p.path.split('.')[0] : null;

    if (isDotPath && groupKey) {
      // Ref child — add to ref group (group label = first segment)
      if (!colGroupsMap.has(groupKey)) {
        colGroupsMap.set(groupKey, { label: groupKey, options: [] });
      }
      // Label inside group: just the tail (after →)
      const shortLabel = p.label.includes(' → ') ? p.label.split(' → ').slice(1).join(' → ') : p.label;
      colGroupsMap.get(groupKey)!.options.push({ value: p.path, label: shortLabel });
    } else if (p.type === 'date' || p.type === 'datetime') {
      // Date column — its own group with feature sub-options
      const dateGroup: ColGroup = { label: p.label, options: [{ value: p.path, label: 'Raw value' }] };
      const features = p.type === 'date'
        ? DATE_FEATURES.filter(f => f.value !== 'hour')
        : DATE_FEATURES;
      for (const f of features) {
        dateGroup.options.push({ value: `${p.path}:${f.value}`, label: f.label });
      }
      colGroupsMap.set(p.path, dateGroup);
    } else {
      directGroup.options.push({ value: p.path, label: p.label });
    }
  }

  // Final grouped list: direct first, then ref/date groups
  const colOptions: ColGroup[] = [];
  if (directGroup.options.length > 0) colOptions.push(directGroup);
  for (const g of colGroupsMap.values()) colOptions.push(g);

  // Flat list for lookups (finding currently selected option)
  const colOptionsFlat: ColOption[] = colOptions.flatMap(g => g.options);

  // Filtered grouped options for Y column (exclude reference/image/bool when aggregating)
  const yColOptions: ColGroup[] = draft.aggregate === 'none'
    ? colOptions
    : colOptions.map(g => ({ ...g, options: g.options.filter(o => numericPathSet.has(o.value)) })).filter(g => g.options.length > 0);

  const set = <K extends keyof ChartConfig>(key: K, val: ChartConfig[K]) =>
    setDraft(d => ({ ...d, [key]: val }));

  return (
    <div
      style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 3000 }}
      onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div style={{ background: 'var(--color-surface)', border: '1px solid var(--color-border)', borderRadius: 12, width: 'min(440px, 94vw)', display: 'flex', flexDirection: 'column', boxShadow: '0 12px 40px rgba(0,0,0,0.22)', color: 'var(--color-text)', overflow: 'hidden' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '14px 20px', borderBottom: '1px solid var(--color-border)' }}>
          <span style={{ fontWeight: 600, fontSize: 15 }}>{isNew ? 'Add Chart' : 'Edit Chart'}</span>
          <button onClick={onClose} className="app-dialog-close" aria-label="Close">×</button>
        </div>
        <div style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: 12 }}>
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
                menuPlacement="auto"
              />
            </div>
          </div>
          <div style={{ display: 'flex', gap: 12 }}>
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
              <label className="app-dialog-label" style={{ marginBottom: 0 }}>X column</label>
              <Select
                styles={dialogSelectStyles}
                value={colOptionsFlat.find(o => o.value === draft.xColumn) ?? null}
                options={colOptions}
                onChange={opt => set('xColumn', opt?.value ?? '')}
                placeholder="— select —"
                isClearable
                menuPlacement="auto"
              />
            </div>
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
              <label className="app-dialog-label" style={{ marginBottom: 0 }}>Aggregate</label>
              <Select
                styles={dialogSelectStyles}
                isSearchable={false}
                value={AGGREGATE_OPTIONS.find(o => o.value === draft.aggregate) ?? null}
                options={AGGREGATE_OPTIONS}
                onChange={opt => set('aggregate', (opt?.value ?? 'count') as AggregateFunc)}
                menuPlacement="auto"
              />
            </div>
          </div>
          {needsYCol && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              <label className="app-dialog-label" style={{ marginBottom: 0 }}>Y column</label>
              <Select
                styles={dialogSelectStyles}
                value={colOptionsFlat.find(o => o.value === draft.yColumn) ?? null}
                options={yColOptions}
                onChange={opt => set('yColumn', opt?.value ?? '')}
                placeholder="— select —"
                isClearable
                menuPlacement="auto"
              />
            </div>
          )}
          {hasGroupBy && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              <label className="app-dialog-label" style={{ marginBottom: 0 }}>
                Group by <span style={{ fontWeight: 400, color: 'var(--color-text-muted)' }}>(optional)</span>
              </label>
              <Select
                styles={dialogSelectStyles}
                value={draft.groupBy ? colOptionsFlat.find(o => o.value === draft.groupBy) ?? null : null}
                options={colOptions}
                onChange={opt => set('groupBy', opt?.value || undefined)}
                placeholder="— none —"
                isClearable
                menuPlacement="auto"
              />
            </div>
          )}
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
      aggregate: 'count',
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
          onSave={handleSaveChart}
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
            isDraggable={canEdit}
            isResizable={canEdit}
          >
            {charts.map(chart => {
              const rows = state.getRows(chart.table);
              const { data, seriesKeys } = aggregateData(rows, chart.xColumn, chart.yColumn, chart.aggregate, chart.groupBy, state.resolveColumnPath, chart.table);
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
                      cursor: canEdit ? 'grab' : 'default',
                      flexShrink: 0,
                      userSelect: 'none',
                    }}
                  >
                    <span style={{ fontWeight: 600, fontSize: 13, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {chart.title}
                    </span>
                    {canEdit && (
                      <div style={{ display: 'flex', gap: 4, flexShrink: 0 }} onMouseDown={e => e.stopPropagation()}>
                        <button
                          onClick={() => { setEditingChart(chart); setIsNewChart(false); }}
                          style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--color-text-muted)', padding: '2px 6px', fontSize: 12, borderRadius: 4 }}
                        >Edit</button>
                        <button
                          onClick={() => handleDeleteChart(chart.id)}
                          style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--color-danger)', padding: '2px 6px', fontSize: 12, borderRadius: 4 }}
                        >×</button>
                      </div>
                    )}
                  </div>
                  <div style={{ flex: 1, minHeight: 0, padding: '8px 4px 4px' }}>
                    <ChartRenderer config={chart} data={data} seriesKeys={seriesKeys} />
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
