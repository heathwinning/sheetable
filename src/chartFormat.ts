// chartFormat.ts
// Utilities for value calculation (expr-eval) and display formatting (d3-format / d3-time-format).
//
// valueCalc: an expr-eval expression. Variable `value` holds the raw number; row column names
//   are also available as variables. e.g. "value / 1000"  or  "round(distance / time, 2)"
//
// valueFormat: a template string using {variable:formatSpec} syntax.
//   Numbers use d3-format specs  — e.g. "{value:,.2f} km"  →  "1,234.56 km"
//   Dates use d3-time-format specs — e.g. "{date:%b %Y}"  →  "May 2026"
//   Plain {variable} substitutions are also supported.

import { Parser } from 'expr-eval';
import { format as d3Format } from 'd3-format';
import { timeFormat } from 'd3-time-format';
import type { ColumnModifier } from './types';

// ── Caches ────────────────────────────────────────────────────────────────────

const calcCache = new Map<string, ReturnType<typeof buildCalc>>();
const numFormatCache = new Map<string, (n: number) => string>();
const dateFormatCache = new Map<string, (d: Date) => string>();
const parser = new Parser();

function buildCalc(expr: string) {
  // Support {colName} syntax — strip braces so expr-eval sees bare identifiers
  const normalized = expr.replace(/\{([^}]+)\}/g, (_, n: string) => n).trim();
  try {
    const compiled = parser.parse(normalized);
    return (v: number, ctx?: Record<string, number>) => {
      try { return Number(compiled.evaluate({ value: v, ...ctx })); }
      catch { return v; }
    };
  } catch {
    return (v: number) => v;
  }
}

export function getCalc(expr: string) {
  if (!calcCache.has(expr)) calcCache.set(expr, buildCalc(expr));
  return calcCache.get(expr)!;
}

function getNumFormat(spec: string): (n: number) => string {
  if (!numFormatCache.has(spec)) {
    try { numFormatCache.set(spec, d3Format(spec)); }
    catch { numFormatCache.set(spec, n => String(n)); }
  }
  return numFormatCache.get(spec)!;
}

function getDateFormat(spec: string): (d: Date) => string {
  if (!dateFormatCache.has(spec)) dateFormatCache.set(spec, timeFormat(spec));
  return dateFormatCache.get(spec)!;
}

// ── Template renderer ─────────────────────────────────────────────────────────
// Syntax: {varName:formatSpec} or {varName}
// Format spec starting with % → d3-time-format (date).
// Anything else → d3-format (number).

function applyTemplate(tpl: string, ctx: Record<string, unknown>): string {
  return tpl.replace(/\{(\w+)(?::([^}]*))?\}/g, (_, name: string, spec: string | undefined) => {
    const val = ctx[name];
    if (val === undefined || val === null) return '';
    if (!spec) return String(val);

    if (spec.startsWith('%')) {
      const d = val instanceof Date ? val : new Date(String(val));
      if (isNaN(d.getTime())) return String(val);
      return getDateFormat(spec)(d);
    }

    const num = Number(val);
    if (isNaN(num)) return String(val);
    return getNumFormat(spec)(num);
  });
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
 * @param rawValue  The raw aggregated number.
 * @param opts      valueCalc (expr-eval expression), valueFormat (template string), yModifier (legacy).
 * @param rowCtx    Optional full row context for multi-column expressions.
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

  if (!valueCalc && !valueFormat) {
    if (yModifier) return legacyFormat(rawValue, yModifier);
    return Number.isInteger(rawValue) ? String(rawValue) : rawValue.toFixed(2);
  }

  const calculated = valueCalc ? getCalc(valueCalc)(rawValue, rowCtx) : rawValue;

  if (valueFormat) {
    const ctx: Record<string, unknown> = { value: calculated, ...rowCtx };
    try { return applyTemplate(valueFormat, ctx); }
    catch { return String(calculated); }
  }

  return Number.isInteger(calculated) ? String(calculated) : calculated.toFixed(2);
}
