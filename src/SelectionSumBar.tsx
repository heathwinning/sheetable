import React from 'react';
import type { GridApi, CellRange } from 'ag-grid-community';

export interface SelectionStats {
  sum: number;
  count: number;
  avg: number;
}

/**
 * Reads numeric values from all ag-grid cell ranges and returns stats.
 * Returns null if no numeric cells are selected.
 */
export function computeCellSelectionStats(api: GridApi): SelectionStats | null {
  const ranges: CellRange[] | null = api.getCellRanges();
  if (!ranges || ranges.length === 0) return null;

  const nums: number[] = [];
  for (const range of ranges) {
    const startIdx = range.startRow?.rowIndex ?? 0;
    const endIdx = range.endRow?.rowIndex ?? 0;
    const minRow = Math.min(startIdx, endIdx);
    const maxRow = Math.max(startIdx, endIdx);
    for (let i = minRow; i <= maxRow; i++) {
      const rowNode = api.getDisplayedRowAtIndex(i);
      if (!rowNode || rowNode.rowPinned) continue;
      for (const col of range.columns) {
        const raw = api.getCellValue({ rowNode, colKey: col });
        if (raw === null || raw === undefined || raw === '') continue;
        const n = Number(raw);
        if (!Number.isNaN(n)) nums.push(n);
      }
    }
  }

  if (nums.length === 0) return null;
  const sum = nums.reduce((a, b) => a + b, 0);
  return { sum, count: nums.length, avg: sum / nums.length };
}

interface SelectionSumBarProps {
  stats: SelectionStats | null;
  /** 'fixed' (default) — bottom-right of viewport. 'above' — absolute, just above parent's bottom edge. */
  placement?: 'fixed' | 'above';
}

function fmt(n: number): string {
  // Use compact notation for large numbers, otherwise show up to 6 significant digits
  if (Math.abs(n) >= 1e9) return n.toLocaleString(undefined, { maximumFractionDigits: 2 });
  const s = parseFloat(n.toPrecision(9));
  return s.toLocaleString(undefined, { maximumFractionDigits: 6 });
}

export const SelectionSumBar: React.FC<SelectionSumBarProps> = ({ stats, placement = 'fixed' }) => {
  if (!stats) return null;
  return (
    <div className={`selection-sum-bar${placement === 'above' ? ' selection-sum-bar--above' : ''}`}>
      {stats.count > 1 && (
        <span className="selection-sum-item">
          <span className="selection-sum-label">Avg</span>
          {fmt(stats.avg)}
        </span>
      )}
      <span className="selection-sum-item">
        <span className="selection-sum-label">Count</span>
        {stats.count}
      </span>
      <span className="selection-sum-item selection-sum-total">
        <span className="selection-sum-label">Sum</span>
        {fmt(stats.sum)}
      </span>
    </div>
  );
};
