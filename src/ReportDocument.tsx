import { createContext, useContext, useState } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { createUsePuck, FieldLabel, Puck, Render } from '@puckeditor/core';
import type { Config, ConfigParams, Field } from '@puckeditor/core';
import type { ReportBlocks, ReportData, ReportRoot } from '../shared/reportTypes';
import * as api from './api';
import type { UseAppStateReturn } from './useAppState';
import { useSheetRows } from './useSheetRows';
import { Thumb } from './Thumb';
import { ReportChart } from './ChartSheetPage';
import { reportThemes } from './reportTemplates';
import { INTERNAL_ROW_ID } from './types';
import type { ChartConfig } from './types';
import '@puckeditor/core/puck.css';
import './ReportDocument.css';

const ReportContext = createContext<{ bookId: string; state: UseAppStateReturn } | null>(null);

function useReport() {
  const report = useContext(ReportContext);
  if (!report) throw new Error('Missing report context');
  return report;
}

function WorkbookImage({ image, alt, className, focus }: { image: string; alt: string; className?: string; focus?: string }) {
  const { bookId } = useReport();
  return <img className={className} src={api.imageUrl(bookId, image)} alt={alt} style={{ objectPosition: focus }} />;
}

function ImageField({ value, onChange, label }: { value: string; onChange: (value: string) => void; label: string }) {
  const { bookId } = useReport();
  const [assets, setAssets] = useState<string[]>([]);
  const [cursor, setCursor] = useState<string>();
  const [browsing, setBrowsing] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');

  async function browse(next?: string) {
    setBusy(true);
    setError('');
    try {
      const result = await api.listReportAssets(bookId, next);
      setAssets(previous => next ? [...previous, ...result.images] : result.images);
      setCursor(result.cursor);
      setBrowsing(true);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Could not load workbook images.');
    } finally {
      setBusy(false);
    }
  }

  async function upload(file?: File) {
    if (!file) return;
    setBusy(true);
    setError('');
    try {
      const { key, uploadUrl } = await api.getUploadUrl(bookId, file.name);
      await api.uploadImage(uploadUrl, file);
      onChange(key);
      setBrowsing(false);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Could not upload the image.');
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="report-image-field">
      <FieldLabel label={label}>
        <input type="file" aria-label={`Upload ${label.toLowerCase()}`}
          accept="image/jpeg,image/png,image/webp,image/gif,image/avif" disabled={busy}
          onChange={event => {
            void upload(event.currentTarget.files?.[0]);
            event.currentTarget.value = '';
          }} />
      </FieldLabel>
      {value && <WorkbookImage image={value} alt={`Selected ${label.toLowerCase()}`} className="report-image-selected" />}
      <div className="report-field-actions">
        <button type="button" disabled={busy} onClick={() => browsing ? setBrowsing(false) : void browse()}>
          {browsing ? 'Close library' : 'Workbook images'}
        </button>
        {value && <button type="button" disabled={busy} onClick={() => onChange('')}>Remove</button>}
      </div>
      {busy && <span role="status">Working…</span>}
      {error && <span role="alert">{error}</span>}
      {browsing && (
        <>
          <div className="report-image-picker">
            {assets.map((asset, index) => (
              <button type="button" key={asset} aria-label={`Choose workbook image ${index + 1}`}
                onClick={() => { onChange(asset); setBrowsing(false); }}>
                <Thumb bookId={bookId} imageKey={asset} />
              </button>
            ))}
          </div>
          {assets.length === 0 && <small>This workbook has no uploaded images yet.</small>}
          {cursor && <button type="button" disabled={busy} onClick={() => void browse(cursor)}>More images</button>}
        </>
      )}
    </div>
  );
}

const getChartKey = (sheet: string, id: string) => JSON.stringify([sheet, id]);

function ChartField({ value, onChange }: { value: string; onChange: (value: string) => void }) {
  const { state } = useReport();
  const charts = state.chartSheetIds.flatMap(sheet => (state.getChartSheet(sheet)?.charts ?? [])
    .map(chart => ({
      key: getChartKey(sheet, chart.id),
      name: `${sheet} · ${chart.title || (chart.type === 'table' ? 'Pivot table' : chart.type)}`,
    })));
  return (
    <FieldLabel label="Workbook chart">
      <select aria-label="Workbook chart" value={value} onChange={event => onChange(event.target.value)}>
        <option value="">{charts.length ? 'Choose a chart' : 'No charts in this workbook'}</option>
        {value && !charts.some(chart => chart.key === value) && <option value={value}>Missing chart</option>}
        {charts.map(chart => <option key={chart.key} value={chart.key}>{chart.name}</option>)}
      </select>
    </FieldLabel>
  );
}

function TableField({ value, onChange }: { value: string; onChange: (value: string) => void }) {
  const { state } = useReport();
  return (
    <FieldLabel label="Workbook table">
      <select value={value} onChange={event => onChange(event.target.value)}>
        <option value="">Choose a table</option>
        {value && !state.tableIds.includes(value) && <option value={value}>{value} (missing)</option>}
        {state.tableIds.map(table => <option key={table} value={table}>{table}</option>)}
      </select>
    </FieldLabel>
  );
}

function ReportTable({ table, title }: { table: string; title: string }) {
  const { state } = useReport();
  const result = useSheetRows(state, table ? [table] : []);
  const schema = state.getSchema(table);
  if (result.error) return <p data-report-data="error" role="alert">{result.error} <button type="button" onClick={result.retry}>Retry</button></p>;
  if (result.loading) return <p data-report-data="loading" role="status">Loading {table}…</p>;
  if (!schema) return <p className="report-placeholder">{table ? `Table “${table}” is unavailable.` : 'Choose a workbook table.'}</p>;
  const columns = schema.columns.filter(column => column.showInGrid !== false && column.type !== 'image'
    && !(column.type === 'list' && column.listOf === 'image'));
  const rows = state.getRows(table);
  return (
    <section className="report-table-section">
      {title && <h2>{title}</h2>}
      {rows.length === 0 || columns.length === 0 ? <p>No rows to display.</p> : (
        <table>
          <thead><tr>{columns.map(column => <th key={column.name} scope="col">{column.displayName ?? column.name}</th>)}</tr></thead>
          <tbody>
            {rows.map((row, index) => (
              <tr key={row[INTERNAL_ROW_ID] || index}>
                {columns.map(column => (
                  <td key={column.name}>{state.resolveColumnPath(table, row, column.name)}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </section>
  );
}

function ChartBlock({ chart: key, title }: { chart: string; title: string }) {
  const { state } = useReport();
  const charts = state.chartSheetIds.flatMap(sheet => (state.getChartSheet(sheet)?.charts ?? [])
    .map(chart => ({ sheet, chart })));
  const exact = charts.find(({ sheet, chart }) => getChartKey(sheet, chart.id) === key);
  let chart: ChartConfig | undefined = exact?.chart;
  if (!chart && key) {
    // Chart IDs survive sheet renames. Fall back only when the ID is unambiguous.
    const matches = charts.filter(({ chart }) => key.endsWith(`,${JSON.stringify(chart.id)}]`));
    if (matches.length === 1) chart = matches[0].chart;
  }
  const hasTable = chart && state.tableIds.includes(chart.table);
  const result = useSheetRows(state, chart && hasTable ? [chart.table] : []);
  if (result.error) return <p data-report-data="error" role="alert">{result.error} <button type="button" onClick={result.retry}>Retry</button></p>;
  if (result.loading) return <p data-report-data="loading" role="status">Loading chart…</p>;
  if (!chart || !hasTable) return <p className="report-placeholder">{key ? 'Chart data is unavailable.' : 'Choose a workbook chart.'}</p>;
  const isPivot = chart.type === 'table';
  return (
    <section className={isPivot ? 'report-table-section' : 'report-chart-block'}>
      {(title || chart.title) && <h2>{title || chart.title}</h2>}
      <div className={isPivot ? 'report-pivot' : 'report-chart'}>
        <ReportChart config={chart} state={state} />
      </div>
    </section>
  );
}

function ReportPaper({ children, theme, accent, typography, spacing, footer }: {
  children: ReactNode; theme: string; accent: string; typography: string; spacing: string; footer: string;
}) {
  const style = /^#[0-9a-f]{6}$/i.test(accent) ? { '--report-accent': accent } as CSSProperties : undefined;
  return (
    <article className="report-paper" data-theme={theme} data-typography={typography} data-spacing={spacing} style={style}>
      <div className="report-content">{children}</div>
      {footer && <footer className="report-footer">{footer}</footer>}
    </article>
  );
}

type ReportConfig = Config<ConfigParams<ReportBlocks, ReportRoot>>;
const useReportPuck = createUsePuck<ReportConfig>();

function AccentField({ value, onChange }: { value: string; onChange: (value: string) => void }) {
  const theme = useReportPuck(({ appState }) => appState.data.root.props?.theme);
  const baseAccent = reportThemes.find(option => option.value === theme)?.accent ?? reportThemes[0].accent;
  return (
    <div className="report-field-actions">
      <FieldLabel label="Accent colour">
        <input type="color" value={value || baseAccent} onChange={event => onChange(event.target.value)} />
      </FieldLabel>
      {value && <button type="button" onClick={() => onChange('')}>Use theme colour</button>}
    </div>
  );
}

const imageField = (label: string): Field => ({
  type: 'custom', label,
  render: ({ value, onChange }) => <ImageField label={label} value={value ?? ''} onChange={onChange} />,
});

const focus: Field = {
  type: 'select', label: 'Crop focus',
  options: [
    { label: 'Centre', value: 'center' }, { label: 'Top', value: 'top' }, { label: 'Bottom', value: 'bottom' },
    { label: 'Left', value: 'left' }, { label: 'Right', value: 'right' },
  ],
};

const config: ReportConfig = {
  categories: {
    editorial: { title: 'Page design', components: ['Cover', 'Text', 'Image', 'Callout', 'PageBreak'] },
    workbook: { title: 'Workbook data', components: ['Table', 'Chart'] },
  },
  root: {
    defaultProps: { theme: 'editorial', accent: '', typography: 'theme', spacing: 'airy', footer: '' },
    fields: {
      theme: { type: 'select', options: reportThemes, label: 'Theme' },
      accent: {
        type: 'custom', label: 'Accent colour',
        render: ({ value, onChange }) => <AccentField value={value} onChange={onChange} />,
      },
      typography: {
        type: 'select', label: 'Heading font', options: [
          { label: 'Theme default', value: 'theme' }, { label: 'Serif', value: 'serif' }, { label: 'Sans serif', value: 'sans' },
        ],
      },
      spacing: { type: 'select', label: 'Spacing', options: [{ label: 'Airy', value: 'airy' }, { label: 'Compact', value: 'compact' }] },
      footer: { type: 'text', label: 'Footer' },
    },
    render: ({ children, theme, accent, typography, spacing, footer }) =>
      <ReportPaper theme={theme} accent={accent} typography={typography} spacing={spacing} footer={footer}>{children}</ReportPaper>,
  },
  components: {
    Cover: {
      label: 'Cover',
      defaultProps: { title: 'Report title', subtitle: '', label: '', logo: '', image: '', alt: '', layout: 'editorial', focus: 'center' },
      fields: {
        title: { type: 'text', label: 'Title', contentEditable: true },
        subtitle: { type: 'text', label: 'Subtitle', contentEditable: true },
        label: { type: 'text', label: 'Organisation or edition' },
        layout: {
          type: 'select', label: 'Layout', options: [
            { label: 'Editorial', value: 'editorial' }, { label: 'Photo overlay', value: 'overlay' },
          ],
        },
        logo: imageField('Logo'), image: imageField('Cover photo'),
        alt: { type: 'text', label: 'Photo alt text' }, focus,
      },
      render: ({ title, subtitle, label, logo, image, alt, layout, focus, puck }) => (
        <header className="report-cover" data-layout={layout} data-photo={Boolean(image)}>
          {image && <WorkbookImage className="report-cover-photo" image={image} alt={alt} focus={focus} />}
          {!image && puck.isEditing && layout !== 'overlay' && <div className="report-placeholder">Add a cover photo</div>}
          <div className="report-cover-copy">
            {(logo || label) && (
              <div className="report-cover-brand">
                {logo && <WorkbookImage image={logo} alt={label ? '' : 'Organisation logo'} />}
                {label && <span>{label}</span>}
              </div>
            )}
            <h1>{title}</h1>
            {subtitle && <div className="report-cover-subtitle">{subtitle}</div>}
          </div>
        </header>
      ),
    },
    Text: {
      label: 'Text section',
      defaultProps: { heading: 'Section heading', body: 'Write your report here.', layout: 'wide' },
      fields: {
        heading: { type: 'text', label: 'Heading', contentEditable: true },
        body: { type: 'textarea', label: 'Text', contentEditable: true },
        layout: { type: 'select', label: 'Layout', options: [{ label: 'Stacked', value: 'wide' }, { label: 'Split heading', value: 'split' }] },
      },
      render: ({ heading, body, layout }) => (
        <section className="report-text-section" data-layout={layout}>
          {heading && <h2>{heading}</h2>}
          <div className="report-text">{body}</div>
        </section>
      ),
    },
    Image: {
      label: 'Photo and caption',
      defaultProps: { image: '', alt: '', caption: '', size: 'landscape', focus: 'center' },
      fields: {
        image: imageField('Photo'), alt: { type: 'text', label: 'Alt text' },
        caption: { type: 'text', label: 'Caption' },
        size: {
          type: 'select', label: 'Frame', options: [
            { label: 'Landscape', value: 'landscape' }, { label: 'Square', value: 'square' }, { label: 'Original', value: 'original' },
          ],
        },
        focus,
      },
      render: ({ image, alt, caption, size, focus }) => (
        <figure className="report-figure" data-size={size}>
          {image ? <WorkbookImage image={image} alt={alt} focus={focus} /> : <div className="report-placeholder">Choose a photo</div>}
          {caption && <figcaption>{caption}</figcaption>}
        </figure>
      ),
    },
    Callout: {
      defaultProps: { body: 'Add a finding or a key takeaway.' },
      fields: { body: { type: 'textarea', label: 'Text', contentEditable: true } },
      render: ({ body }) => <aside className="report-callout report-text">{body}</aside>,
    },
    Table: {
      label: 'Workbook table',
      defaultProps: { title: '', table: '' },
      fields: {
        title: { type: 'text', label: 'Heading' },
        table: { type: 'custom', label: 'Workbook table', render: ({ value, onChange }) => <TableField value={value} onChange={onChange} /> },
      },
      render: ({ table, title }) => <ReportTable table={table} title={title} />,
    },
    Chart: {
      label: 'Workbook chart',
      defaultProps: { title: '', chart: '' },
      fields: {
        title: { type: 'text', label: 'Override heading' },
        chart: { type: 'custom', label: 'Workbook chart', render: ({ value, onChange }) => <ChartField value={value} onChange={onChange} /> },
      },
      render: ({ chart, title }) => <ChartBlock chart={chart} title={title} />,
    },
    PageBreak: {
      label: 'Page break',
      defaultProps: {},
      render: ({ puck }) => <div className="report-page-break">{puck.isEditing ? 'Page break' : null}</div>,
    },
  },
};

const editorOverrides = { headerActions: () => <></> };

export function ReportDocument({ state, bookId, data, editing, onChange }: {
  state: UseAppStateReturn;
  bookId: string;
  data: ReportData;
  editing?: boolean;
  onChange?: (data: ReportData) => void;
}) {
  return (
    <ReportContext.Provider value={{ state, bookId }}>
      {editing ? (
        <Puck config={config} data={data} onChange={onChange} overrides={editorOverrides}
          headerTitle="Report designer"
          viewports={[{ width: 794, height: 'auto', label: 'A4' }, { width: 390, height: 'auto', label: 'Mobile' }]} />
      ) : <Render config={config} data={data} />}
    </ReportContext.Provider>
  );
}
