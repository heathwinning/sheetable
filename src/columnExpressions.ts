import type { DateFeature } from './types';

export interface ColumnPath {
  path: string;
  label: string;
  type?: string;
}

export const DATE_FEATURES: { value: DateFeature; label: string; description?: string }[] = [
  { value: 'year', label: 'Year' },
  { value: 'quarter', label: 'Quarter' },
  { value: 'yearmonth', label: 'Year-Month' },
  { value: 'month', label: 'Month name' },
  { value: 'monthnum', label: 'Month number' },
  { value: 'week', label: 'Week of year', description: 'Sunday-start calendar weeks. Week 1 contains January 1.' },
  { value: 'weekofmonth', label: 'Week of month', description: 'Sunday-start calendar weeks (1–6). Week 1 contains the first day of the month.' },
  { value: 'dayofweek', label: 'Day of week' },
  { value: 'day', label: 'Day of month' },
  { value: 'hour', label: 'Hour' },
];

export function getColumnTransformations(type?: string) {
  if (type !== 'date' && type !== 'datetime') return [];
  return type === 'date' ? DATE_FEATURES.filter(feature => feature.value !== 'hour') : DATE_FEATURES;
}

export function parseColumnExpression(expression: string): { col: string; feature: DateFeature | null } {
  const separator = expression.lastIndexOf(':');
  const feature = expression.slice(separator + 1);
  if (separator < 0 || !DATE_FEATURES.some(option => option.value === feature)) {
    return { col: expression, feature: null };
  }
  return { col: expression.slice(0, separator), feature: feature as DateFeature };
}

export function stripFeature(expression: string): string {
  return parseColumnExpression(expression).col;
}

export function getFeature(expression: string): DateFeature | undefined {
  return parseColumnExpression(expression).feature ?? undefined;
}

export function formatColumnExpression(expression: string, columns: ColumnPath[]): string {
  const { col, feature } = parseColumnExpression(expression);
  const label = columns.find(column => column.path === col)?.label ?? col;
  return feature ? `${label} · ${DATE_FEATURES.find(option => option.value === feature)!.label}` : label;
}

const MONTH_NAME_FORMAT = new Intl.DateTimeFormat(undefined, { month: 'long' });
const WEEKDAY_NAME_FORMAT = new Intl.DateTimeFormat(undefined, { weekday: 'long' });

export function applyDateFeature(raw: string, feature: DateFeature): string {
  if (!raw.trim()) return '';
  const date = new Date(raw);
  if (isNaN(date.getTime())) return raw;
  switch (feature) {
    case 'year': return String(date.getFullYear());
    case 'quarter': return `Q${Math.ceil((date.getMonth() + 1) / 3)}`;
    case 'yearmonth': return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-01`;
    case 'month': return MONTH_NAME_FORMAT.format(date);
    case 'monthnum': return String(date.getMonth() + 1).padStart(2, '0');
    case 'week': {
      const januaryFirst = new Date(date.getFullYear(), 0, 1);
      const dayOfYear = (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - Date.UTC(date.getFullYear(), 0, 1)) / 86400000 + 1;
      return String(Math.ceil((dayOfYear + januaryFirst.getDay()) / 7)).padStart(2, '0');
    }
    case 'weekofmonth': {
      const firstDay = new Date(date.getFullYear(), date.getMonth(), 1).getDay();
      return String(Math.ceil((date.getDate() + firstDay) / 7)).padStart(2, '0');
    }
    case 'dayofweek': return WEEKDAY_NAME_FORMAT.format(date);
    case 'day': return String(date.getDate()).padStart(2, '0');
    case 'hour': return String(date.getHours()).padStart(2, '0') + ':00';
  }
}

export function applyColumnExpression(raw: string, expression: string): string {
  const { feature } = parseColumnExpression(expression);
  return feature ? applyDateFeature(raw, feature) : raw;
}
