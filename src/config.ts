import type { TableSchema } from './types';

// Schema is stored as a JSON file in the Drive folder
export interface ProjectConfig {
  tables: TableSchema[];
}

export function serializeConfig(config: ProjectConfig): string {
  return JSON.stringify(config, null, 2);
}

export function parseConfig(text: string): ProjectConfig {
  return JSON.parse(text) as ProjectConfig;
}
