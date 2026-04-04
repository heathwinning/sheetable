import type { TableSchema, ChartSheet } from './types';

// Schema is stored as a JSON file in the Drive folder
export interface ProjectConfig {
  tables: TableSchema[];
  chartSheets?: ChartSheet[];
}

export interface BooksConfig {
  books: Array<{
    id: string;
    name: string;
  }>;
}

export function serializeConfig(config: ProjectConfig): string {
  return JSON.stringify(config, null, 2);
}

export function parseConfig(text: string): ProjectConfig {
  return JSON.parse(text) as ProjectConfig;
}

export function serializeBooksConfig(config: BooksConfig): string {
  return JSON.stringify(config, null, 2);
}

export function parseBooksConfig(text: string): BooksConfig {
  return JSON.parse(text) as BooksConfig;
}
