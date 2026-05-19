// Shared AG Grid defaultColDef applied to every grid instance in the app.
// Centralised here so trimming (and other global behaviours) only need to be maintained once.
import type { ColDef } from 'ag-grid-community';

export const sharedDefaultColDef: ColDef = {
  // Trim leading/trailing whitespace on every cell edit — prevents mobile keyboards
  // from silently appending spaces.
  valueParser: (params) =>
    typeof params.newValue === 'string' ? params.newValue.trim() : params.newValue,
};
