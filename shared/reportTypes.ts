import type { Data } from '@puckeditor/core';

export type SheetType = 'table' | 'chart' | 'view' | 'report';
export type SheetOrderEntry = { type: SheetType; name: string; hidden?: boolean };

export type ReportSummary = { id: string; name: string };
export type ReportRoot = { theme: string; accent: string; typography: string; spacing: string; footer: string };
export type ReportBlocks = {
  Cover: { title: string; subtitle: string; label: string; logo: string; image: string; alt: string; layout: string; focus: string };
  Text: { heading: string; body: string; layout: string };
  Image: { image: string; alt: string; caption: string; size: string; focus: string };
  Callout: { body: string };
  Table: { title: string; table: string };
  Chart: { title: string; chart: string };
  PageBreak: Record<never, never>;
};
export type ReportData = Data<ReportBlocks, ReportRoot>;
export type ReportDefinition = { version: 1; data: ReportData };
export type ReportSheet = ReportSummary & { revision: number; definition: ReportDefinition };
export type ReportAssetPage = { images: string[]; cursor?: string };

const blockTypes = new Set(['Cover', 'Text', 'Image', 'Callout', 'Table', 'Chart', 'PageBreak']);

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function isReportProps(value: unknown): boolean {
  return isRecord(value) && Object.values(value).every(prop => typeof prop === 'string' && prop.length <= 20_000);
}

export function isReportDefinition(value: unknown): value is ReportDefinition {
  if (!isRecord(value) || value.version !== 1 || !isRecord(value.data)) return false;
  const { root, content } = value.data;
  if (!isRecord(root) || !isReportProps(root.props) || !Array.isArray(content) || content.length > 150) return false;
  if (value.data.zones !== undefined && (!isRecord(value.data.zones) || Object.keys(value.data.zones).length > 0)) return false;
  const ids = new Set<string>();
  for (const block of content) {
    if (!isRecord(block) || typeof block.type !== 'string' || !blockTypes.has(block.type)
      || !isRecord(block.props) || !isReportProps(block.props)) return false;
    const id = block.props.id;
    if (typeof id !== 'string' || !id || id.length > 120 || ids.has(id)) return false;
    ids.add(id);
  }
  return JSON.stringify(value).length <= 250_000;
}

export function reportName(value: unknown) {
  if (typeof value !== 'string') return null;
  const name = value.trim();
  return name.length > 0 && name.length <= 100 ? name : null;
}
