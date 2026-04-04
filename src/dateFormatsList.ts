// Centralized date format patterns for migration and import
export const DATE_FORMATS = [
  {
    value: 'yyyy/mm/dd',
    label: 'YYYY/MM/DD (2026/04/03)',
    parse: (s: string) => {
      const m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
      return m ? `${m[1]}/${m[2].padStart(2, '0')}/${m[3].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'yyyy-mm-dd',
    label: 'YYYY-MM-DD (2026-04-03)',
    parse: (s: string) => {
      const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      return m ? `${m[1]}/${m[2].padStart(2, '0')}/${m[3].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'dd/mm/yyyy',
    label: 'DD/MM/YYYY (03/04/2026)',
    parse: (s: string) => {
      const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      return m ? `${m[3]}/${m[2].padStart(2, '0')}/${m[1].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'mm/dd/yyyy',
    label: 'MM/DD/YYYY (04/03/2026)',
    parse: (s: string) => {
      const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      return m ? `${m[3]}/${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'auto',
    label: 'Auto-detect',
    parse: (s: string) => {
      // Try ISO first
      const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (iso) return `${iso[1]}/${iso[2].padStart(2, '0')}/${iso[3].padStart(2, '0')}`;
      // Try Date constructor
      const d = new Date(s);
      if (!isNaN(d.getTime())) {
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${y}/${m}/${day}`;
      }
      return null;
    },
  },
];
