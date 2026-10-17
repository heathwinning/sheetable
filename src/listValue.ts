export function parseListItems(value: string | null | undefined): string[] {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    if (Array.isArray(parsed)) return parsed.map(String);
  } catch {
    // fall through
  }
  return value ? [value] : [];
}

export function serializeListItems(items: string[]): string {
  return items.length > 0 ? JSON.stringify(items) : '';
}

export function isValidListItem(listOf: string | undefined, item: string): boolean {
  if (listOf === 'integer') return /^-?\d+$/.test(item);
  if (listOf === 'decimal') return item !== '' && !isNaN(Number(item));
  if (listOf === 'bool') return ['true', 'false'].includes(item.toLowerCase());
  return true;
}
