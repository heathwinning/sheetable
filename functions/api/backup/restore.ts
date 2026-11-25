import type { Env, RequestData } from '../../lib';
import { json, error, requireUser } from '../../lib';
import type { BackupPayload } from '../../../shared/backupTypes';
import { isReportDefinition, reportName } from '../../../shared/reportTypes';
import { restoreBackupBook, validateBackupColumns } from '../../restoreBackup';

export const onRequestPost: PagesFunction<Env, string, RequestData> = async (context) => {
  const user = requireUser(context.data);

  let body: BackupPayload;
  try {
    body = await context.request.json() as BackupPayload;
  } catch {
    return error('Invalid JSON body', 400);
  }

  if (typeof body?.bookName !== 'string' || !body.bookName.trim()) return error('bookName is required');
  if (body.reports !== undefined && (!Array.isArray(body.reports) || body.reports.some(report =>
    !report || typeof report.id !== 'string' || !reportName(report.name) || !isReportDefinition(report.definition)
  ) || new Set(body.reports.map(report => report.id)).size !== body.reports.length)) {
    return error('Invalid reports in backup');
  }
  if (!Array.isArray(body.tables) || (!body.tables.length && !body.reports?.length)) {
    return error('A backup must include tables or reports');
  }
  for (const table of body.tables) {
    const validationError = validateBackupColumns(table.schema?.columns ?? []);
    if (validationError) return error(`${validationError} in table "${table.name}"`);
  }

  return json(await restoreBackupBook(context.env.DB, user.id, body), 201);
};
