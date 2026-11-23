import type { Env, RequestData } from '../../lib';
import { json, error, requireUser } from '../../lib';
import type { BackupPayload } from '../../../shared/backupTypes';
import { restoreBackupBook, validateBackupColumns } from '../../restoreBackup';

export const onRequestPost: PagesFunction<Env, string, RequestData> = async (context) => {
  const user = requireUser(context.data);

  let body: BackupPayload;
  try {
    body = await context.request.json() as BackupPayload;
  } catch {
    return error('Invalid JSON body', 400);
  }

  if (!body.bookName?.trim()) return error('bookName is required');
  if (!body.tables?.length) return error('tables array is required');
  for (const table of body.tables) {
    const validationError = validateBackupColumns(table.schema?.columns ?? []);
    if (validationError) return error(`${validationError} in table "${table.name}"`);
  }

  return json(await restoreBackupBook(context.env.DB, user.id, body), 201);
};
