import React, { useMemo, useRef, useCallback, useEffect } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import { GraphicWalker, GraphicRenderer } from '@kanaries/graphic-walker';
import '@kanaries/graphic-walker/dist/style.css';
import type { VizSpecStore, IMutField, IChart } from '@kanaries/graphic-walker';
import type { UseAppStateReturn } from './useAppState';
import type { ColumnType } from './types';
import { parseTemporalUnknown } from './dateFormat';

const bookPrefix = (bookName?: string) => (bookName ? `/book/${encodeURIComponent(bookName)}` : '');

/** Map our column types to Graphic Walker semantic/analytic types */
function mapColumnType(type: ColumnType): { semanticType: IMutField['semanticType']; analyticType: IMutField['analyticType'] } {
  switch (type) {
    case 'integer':
    case 'decimal':
      return { semanticType: 'quantitative', analyticType: 'measure' };
    case 'date':
    case 'datetime':
      return { semanticType: 'temporal', analyticType: 'dimension' };
    case 'bool':
      return { semanticType: 'nominal', analyticType: 'dimension' };
    default:
      return { semanticType: 'nominal', analyticType: 'dimension' };
  }
}

function normalizeValueForWalker(type: ColumnType, raw: unknown): unknown {
  if (raw == null || raw === '') return null;

  if (type === 'integer' || type === 'decimal') {
    const num = Number(raw);
    return Number.isNaN(num) ? null : num;
  }

  if (type === 'bool') {
    if (raw === true || raw === false) return raw;
    const normalized = String(raw).trim().toLowerCase();
    if (normalized === 'true') return true;
    if (normalized === 'false') return false;
    return null;
  }

  if (type === 'date' || type === 'datetime') {
    const parsed = parseTemporalUnknown(raw);
    if (!parsed) return null;
    if (type === 'date') {
      const y = parsed.getFullYear();
      const m = String(parsed.getMonth() + 1).padStart(2, '0');
      const d = String(parsed.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    return parsed.toISOString();
  }

  return raw;
}

function getColumnTypeFromPath(
  state: UseAppStateReturn,
  tableName: string,
  path: string,
): ColumnType | null {
  const parts = path.split('.').filter(Boolean);
  if (parts.length === 0) return null;

  let currentTable = tableName;
  for (let i = 0; i < parts.length; i++) {
    const schema = state.getSchema(currentTable);
    if (!schema) return null;
    const col = schema.columns.find(c => c.name === parts[i]);
    if (!col) return null;

    if (i === parts.length - 1) {
      return col.type;
    }

    if (col.type !== 'reference' || !col.refTable) {
      return null;
    }
    currentTable = col.refTable;
  }

  return null;
}

function collectLeafColumnPaths(
  state: UseAppStateReturn,
  tableName: string,
  maxDepth = 3,
  stack: string[] = [],
): string[] {
  if (maxDepth <= 0) return [];
  const schema = state.getSchema(tableName);
  if (!schema) return [];

  const out: string[] = [];
  for (const col of schema.columns) {
    if (col.type === 'reference' && col.refTable) {
      // Avoid infinite loops on cyclic references.
      if (stack.includes(col.refTable)) continue;
      const children = collectLeafColumnPaths(
        state,
        col.refTable,
        maxDepth - 1,
        [...stack, tableName],
      );
      for (const child of children) {
        out.push(`${col.name}.${child}`);
      }
    } else {
      out.push(col.name);
    }
  }

  return out;
}

export const ChartSheetPage: React.FC<{ state: UseAppStateReturn }> = ({ state }) => {
  const { bookId, chartId } = useParams<{ bookId?: string; chartId: string }>();
  const navigate = useNavigate();
  const storeRef = useRef<VizSpecStore | null>(null);
  const lastSavedChartsRef = useRef<string>('');

  const chartSheet = chartId ? state.getChartSheet(chartId) : undefined;
  const mode: 'edit' | 'display' = chartSheet?.mode ?? 'edit';
  const selectedTableName = chartSheet?.tableName && state.tableIds.includes(chartSheet.tableName)
    ? chartSheet.tableName
    : state.tableIds[0];

  // Build dataset from only the selected table.
  const { data, fields } = useMemo(() => {
    if (!selectedTableName) return { data: [] as Record<string, unknown>[], fields: [] as IMutField[] };
    const schema = state.getSchema(selectedTableName);
    const rows = state.getRows(selectedTableName);
    if (!schema) return { data: [] as Record<string, unknown>[], fields: [] as IMutField[] };

    const tableFields: IMutField[] = [];
    const referencePathByFid = new Map<string, { sourceCol: string; path: string; pathType: ColumnType }>();

    for (const col of schema.columns) {
      const displayName = col.displayName || col.name;

      if (col.type === 'reference' && col.refTable) {
        // Expose configured and discovered leaf linked paths as first-class chart fields.
        const linkedPaths = Array.from(new Set([
          ...(col.refDisplayColumns ?? []),
          ...(col.refSearchColumns ?? []),
          ...collectLeafColumnPaths(state, col.refTable, 4),
        ]))
          .filter(Boolean)
          .filter((path) => getColumnTypeFromPath(state, col.refTable!, path) !== 'reference');

        for (const path of linkedPaths) {
          const fid = `${col.name}__ref__${path.replace(/\./g, '__')}`;
          const pathType = getColumnTypeFromPath(state, col.refTable, path) ?? 'text';
          const mapped = mapColumnType(pathType);
          const prettyPath = state.model.resolveColumnPathLabel(col.refTable, path);
          const label = `${displayName} → ${prettyPath || path.replace(/\./g, ' → ')}`;

          tableFields.push({
            fid,
            key: fid,
            basename: label,
            name: label,
            semanticType: mapped.semanticType,
            analyticType: mapped.analyticType,
          });

          referencePathByFid.set(fid, { sourceCol: col.name, path, pathType });
        }

        continue;
      }

      const mapped = mapColumnType(col.type);
      tableFields.push({
        fid: col.name,
        key: col.name,
        basename: col.name,
        name: displayName,
        semanticType: mapped.semanticType,
        analyticType: mapped.analyticType,
      });
    }

    const tableData = rows.map((row) => {
      const mappedRow: Record<string, unknown> = {};
      for (const col of schema.columns) {
        if (col.type !== 'reference') {
          mappedRow[col.name] = normalizeValueForWalker(col.type, row[col.name]);
        }
      }

      for (const [fid, refMeta] of referencePathByFid.entries()) {
        const value = state.model.resolveColumnPath(
          selectedTableName,
          row,
          `${refMeta.sourceCol}.${refMeta.path}`,
        );
        mappedRow[fid] = normalizeValueForWalker(refMeta.pathType, value);
      }
      return mappedRow;
    });

    return { data: tableData, fields: tableFields };
  }, [selectedTableName, state.revision]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleSave = useCallback(() => {
    if (!chartId || !storeRef.current) return;
    const charts = storeRef.current.exportCode();
    state.updateChartSheet(chartId, charts as unknown[]);
    lastSavedChartsRef.current = JSON.stringify(charts ?? []);
  }, [chartId, state]);

  const handleToggleMode = useCallback(() => {
    if (!chartId) return;

    // Persist the latest chart layout/spec before entering display mode.
    if (mode === 'edit' && storeRef.current) {
      const charts = storeRef.current.exportCode();
      state.updateChartSheet(chartId, charts as unknown[]);
      lastSavedChartsRef.current = JSON.stringify(charts ?? []);
      state.setChartSheetMode(chartId, 'display');
      return;
    }

    state.setChartSheetMode(chartId, 'edit');
  }, [chartId, mode, state]);

  const handleTableChange = useCallback((tableName: string) => {
    if (!chartId) return;
    state.setChartSheetTable(chartId, tableName);
  }, [chartId, state]);

  const hasUnsavedChartChanges = useCallback((): boolean => {
    const store = storeRef.current;
    if (!store) return false;
    const current = JSON.stringify(store.exportCode() ?? []);
    return current !== lastSavedChartsRef.current;
  }, []);

  // Keep an in-memory snapshot to detect config changes for auto-save.
  useEffect(() => {
    if (!chartSheet) return;
    lastSavedChartsRef.current = JSON.stringify(chartSheet.charts ?? []);
  }, [chartId, chartSheet?.charts]);

  // Auto-save edited charts frequently so users don't have to press Save.
  useEffect(() => {
    if (!chartId || mode !== 'edit') return;

    const timer = window.setInterval(() => {
      const store = storeRef.current;
      if (!store) return;
      const charts = store.exportCode();
      const next = JSON.stringify(charts ?? []);
      if (next === lastSavedChartsRef.current) return;
      state.updateChartSheet(chartId, charts as unknown[]);
      lastSavedChartsRef.current = next;
    }, 1500);

    return () => window.clearInterval(timer);
  }, [chartId, mode, state]);

  // Warn before browser/tab navigation when chart edits haven't been persisted yet.
  useEffect(() => {
    if (mode !== 'edit') return;

    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      if (!hasUnsavedChartChanges()) return;
      e.preventDefault();
      e.returnValue = '';
    };

    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, [hasUnsavedChartChanges, mode]);

  if (!chartId || !chartSheet) {
    return (
      <div className="app-body">
        <div className="main-content">
          <div className="empty-state-main">
            <h2>Chart sheet not found</h2>
            <Link className="btn-secondary" to={bookPrefix(bookId)}>Back</Link>
          </div>
        </div>
      </div>
    );
  }

  if (!selectedTableName) {
    return (
      <div className="app-body chart-sheet-body">
        <div className="chart-sheet-toolbar">
          <button className="btn-secondary btn-sm" onClick={() => navigate(bookPrefix(bookId))}>
            Back
          </button>
        </div>
        <div className="main-content">
          <div className="empty-state-main">
            <h2>No tables available</h2>
            <p>Create a table first, then select it for this chart sheet.</p>
            <button className="btn-secondary" onClick={() => navigate(bookPrefix(bookId))}>Back</button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="app-body chart-sheet-body">
      <div className="chart-sheet-toolbar">
        <label className="chart-table-select-wrap">
          <span>Table</span>
          <select
            className="chart-table-select"
            value={selectedTableName}
            onChange={(e) => handleTableChange(e.target.value)}
          >
            {state.tableIds.map((tableId) => (
              <option key={tableId} value={tableId}>{tableId}</option>
            ))}
          </select>
        </label>
        <button
          className="btn-secondary btn-sm"
          onClick={handleToggleMode}
        >
          {mode === 'edit' ? 'Display Mode' : 'Edit Mode'}
        </button>
        {mode === 'edit' && (
          <button className="btn-primary btn-sm" onClick={handleSave}>
            Save Charts
          </button>
        )}
      </div>
      <div className="chart-sheet-container">
        {mode === 'edit' ? (
          <GraphicWalker
            storeRef={storeRef}
            fields={fields}
            rawFields={fields}
            data={data}
            chart={chartSheet.charts.length > 0 ? chartSheet.charts as IChart[] : undefined}
            appearance="light"
          />
        ) : (
          <GraphicRenderer
            fields={fields}
            rawFields={fields}
            data={data}
            chart={chartSheet.charts.length > 0 ? chartSheet.charts as IChart[] : undefined}
            appearance="light"
          />
        )}
      </div>
    </div>
  );
};
