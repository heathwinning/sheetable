import type { ReportDefinition } from '../shared/reportTypes';

export const reportThemes = [
  { label: 'Editorial', value: 'editorial', accent: '#245b50' },
  { label: 'Studio', value: 'studio', accent: '#514686' },
  { label: 'Journal', value: 'journal', accent: '#983f2a' },
];

export const reportTemplates = [
  { id: 'editorial', name: 'Editorial', description: 'A serif title, a full-width photograph, and room for commentary.' },
  { id: 'studio', name: 'Studio', description: 'A bold cover with a photo behind the title and space for a workbook table.' },
  { id: 'blank', name: 'Blank', description: 'Start with an empty page and add your own sections.' },
] as const;

export type ReportTemplate = typeof reportTemplates[number]['id'];

export function makeReportDefinition(name: string, template: ReportTemplate, table = ''): ReportDefinition {
  const studio = template === 'studio';
  return {
    version: 1,
    data: {
      root: { props: { theme: studio ? 'studio' : 'editorial', accent: '', typography: 'theme', spacing: 'airy', footer: '' } },
      content: template === 'blank' ? [] : [
        {
          type: 'Cover',
          props: {
            id: 'cover', title: name, subtitle: '', label: '', logo: '', image: '', alt: '',
            layout: studio ? 'overlay' : 'editorial', focus: 'center',
          },
        },
        { type: 'Text', props: { id: 'overview', heading: 'Overview', body: '', layout: 'wide' } },
        ...(table ? [{ type: 'Table' as const, props: { id: 'details', title: table, table } }] : []),
      ],
    },
  };
}
