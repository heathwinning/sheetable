// chartFormat.ts
// Utilities for flexible value calculation and formatting in charts.
//
// valueCalc: an expr-eval expression string, variable `value` holds the raw number.
//   e.g. "value * 1000"  or  "round(value / 60, 1)"
//
// valueFormat: a Handlebars template string, `value` holds the calculated number.
//   e.g. "{{value}} km"  or  "{{dateFormat date 'MMM D, YYYY'}}"
//
// Both are optional. If omitted, the raw numeric value is used as-is.

import { Parser } from 'expr-eval';
import Handlebars from 'handlebars';
import dayjs from 'dayjs';
import type { ColumnModifier } from './types';

// ── Handlebars helpers ────────────────────────────────────────────────────────

Handlebars.registerHelper('dateFormat', function (date: unknown, format: unknown) {
  if (!date) return '';
  return dayjs(String(date)).format(typeof format === 'string' ? format : 'YYYY-MM-DD');
});

// ── Caches (avoid re-compiling on every render tick) ─────────────────────────

const calcCache = new Map<string, ReturnType<typeof buildCalc>>();
const templateCache = new Map<string, Handlebars.TemplateDelegate>();
const parser = new Parser();

function buildCalc(expr: string) {
  try {
    const compiled = parser.parse(expr.trim());
    return (v: number, ctx?: Record<string, number>) => {
      try { return Number(compiled.evaluate({ value: v, ...ctx })); }
      catch { return v; }
    };
  } catch {
    return (v: number) => v;
  }
}

function getCalc(expr: string) {
  if (!calcCache.has(expr)) calcCache.set(expr, buildCalc(expr));
  return calcCache.get(expr)!;
}

function getTemplate(tpl: string) {
  if (!templateCache.has(tpl)) templateCache.set(tpl, Handlebars.compile(tpl));
  return templateCache.get(tpl)!;
}

// ── Legacy ColumnModifier fallback ────────────────────────────────────────────

function legacyFormat(n: number, mod: ColumnModifier): string {
  let v = n;
  if (mod.multiplier != null) v = n * mod.multiplier;
  else if (mod.divisor) v = n / mod.divisor;
  const dec = mod.decimals;
  let str = dec !== undefined ? v.toFixed(dec) : (Number.isInteger(v) ? String(v) : v.toFixed(2));
  if (mod.thousands) {
    const parts = str.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    str = parts.join('.');
  }
  return (mod.prefix ?? '') + str + (mod.suffix ?? '');
}

// ── Main exported formatter ───────────────────────────────────────────────────

/**
 * Apply optional calculation + display template to a raw numeric value.
 *
 * Falls back to legacy ColumnModifier if the new fields are absent (backwards compat).
 *
 * @param rawValue  The raw aggregated number from chart data.
 * @param opts      Subset of ChartConfig that carries the formatting fields.
 * @param rowCtx    Optional full row context (numeric values by column name) for multi-column expressions.
 * @returns         A display string.
 */
export function applyChartValueFormat(
  rawValue: number,
  opts: {
    valueCalc?: string;
    valueFormat?: string;
    yModifier?: ColumnModifier;
  },
  rowCtx?: Record<string, number>,
): string {
  const { valueCalc, valueFormat, yModifier } = opts;

  // If neither new field is set, fall back to legacy yModifier
  if (!valueCalc && !valueFormat) {
    if (yModifier) return legacyFormat(rawValue, yModifier);
    // Default: auto number
    return Number.isInteger(rawValue) ? String(rawValue) : rawValue.toFixed(2);
  }

  // 1. Apply calculation
  const calculated = valueCalc ? getCalc(valueCalc)(rawValue, rowCtx) : rawValue;

  // 2. Apply template (or just stringify)
  if (valueFormat) {
    const ctx = { value: Number.isInteger(calculated) ? calculated : parseFloat(calculated.toFixed(4)) };
    try { return String(getTemplate(valueFormat)(ctx)); }
    catch { return String(calculated); }
  }

  return Number.isInteger(calculated) ? String(calculated) : calculated.toFixed(2);
}
