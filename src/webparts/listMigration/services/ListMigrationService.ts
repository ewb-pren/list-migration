import { spfi, SPFx } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PermissionKind } from '@pnp/sp/security';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/attachments';
import '@pnp/sp/files';
import '@pnp/sp/site-users/web';
import '@pnp/sp/security/list';

export interface IListMigrationRequest {
  sourceSiteUrl: string;
  sourceListTitle: string;
  targetSiteUrl: string;
  targetListTitle: string;
}

export interface IListMigrationField {
  InternalName: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
  Required?: boolean;
  DefaultValue?: string;
  SchemaXml: string;
}

export interface IListMigrationItem {
  Id: number;
  Attachments: boolean;
  [name: string]: unknown;
}

export interface IListMigrationPlan {
  key: string;
  sourceListId: string;
  targetListId: string;
  fields: IListMigrationField[];
  items: IListMigrationItem[];
  targetCount: number;
}

export interface IListVerificationDifference {
  side: 'Source' | 'Target';
  itemId: number;
  reason: string;
}

export interface IListVerificationReport {
  sourceRecords: number;
  targetRecords: number;
  matchingRecords: number;
  sourceOnlyRecords: number;
  targetOnlyRecords: number;
  differences: IListVerificationDifference[];
}

interface IMigrationCheckpointEntry {
  targetId: number;
  completed: boolean;
}

interface IMigrationCheckpoint {
  [sourceId: string]: IMigrationCheckpointEntry;
}

const supportedTypes = [
  'Text', 'Note', 'Choice', 'MultiChoice', 'Number', 'Currency',
  'Boolean', 'DateTime', 'URL', 'User', 'UserMulti', 'Guid'
];
const excludedFields = ['ContentType', 'ContentTypeId', 'Attachments'];

function normalizeSiteUrl(value: string, currentSiteUrl: string): string {
  let url: URL;
  try {
    url = new URL(value.trim());
  } catch {
    throw new Error('Enter a valid absolute SharePoint site URL.');
  }
  const current = new URL(currentSiteUrl);
  if (url.protocol !== 'https:' || url.origin.toLowerCase() !== current.origin.toLowerCase() ||
      url.username || url.password || url.search || url.hash) {
    throw new Error('Source and target must be HTTPS sites in the same SharePoint tenant.');
  }
  return url.toString().replace(/\/$/, '');
}

function getCopyFields(source: IListMigrationField[], target: IListMigrationField[]): IListMigrationField[] {
  const fields = source.filter(field => !field.Hidden && !field.ReadOnlyField &&
    excludedFields.indexOf(field.InternalName) < 0);

  for (const field of fields) {
    const destination = target.find(candidate => candidate.InternalName === field.InternalName);
    if (supportedTypes.indexOf(field.TypeAsString) < 0) {
      throw new Error(`Unsupported field: ${field.InternalName} (${field.TypeAsString}).`);
    }
    if (!destination || destination.ReadOnlyField || destination.TypeAsString !== field.TypeAsString) {
      throw new Error(`Missing or incompatible target field: ${field.InternalName}.`);
    }
    if (['Note', 'Choice', 'MultiChoice'].indexOf(field.TypeAsString) >= 0) {
      const sourceXml = new DOMParser().parseFromString(field.SchemaXml, 'application/xml');
      const targetXml = new DOMParser().parseFromString(destination.SchemaXml, 'application/xml');
      const sourceNode = sourceXml.documentElement;
      const targetNode = targetXml.documentElement;
      if (sourceXml.querySelector('parsererror') || targetXml.querySelector('parsererror')) {
        throw new Error(`Invalid field schema: ${field.InternalName}.`);
      }
      if (field.TypeAsString === 'Note') {
        if ([sourceNode, targetNode].some(node => (node.getAttribute('AppendOnly') || '').toUpperCase() === 'TRUE')) {
          throw new Error(`Append-only history needs a migration tool: ${field.InternalName}.`);
        }
        for (const attribute of ['RichText', 'RichTextMode']) {
          if (sourceNode.getAttribute(attribute) !== targetNode.getAttribute(attribute)) {
            throw new Error(`Note setting differs: ${field.InternalName} / ${attribute}.`);
          }
        }
      } else {
        const targetChoices = Array.from(targetXml.querySelectorAll('CHOICE')).map(node => node.textContent);
        const missingChoice = Array.from(sourceXml.querySelectorAll('CHOICE'))
          .some(node => targetChoices.indexOf(node.textContent) < 0);
        if (missingChoice && (targetNode.getAttribute('FillInChoice') || '').toUpperCase() !== 'TRUE') {
          throw new Error(`Target choices are missing: ${field.InternalName}.`);
        }
      }
    }
  }

  for (const field of target) {
    if (field.Required && !field.Hidden && !field.ReadOnlyField && !field.DefaultValue &&
        excludedFields.indexOf(field.InternalName) < 0 &&
        !fields.some(candidate => candidate.InternalName === field.InternalName)) {
      throw new Error(`Required target field has no mapping: ${field.InternalName}.`);
    }
  }
  if (!fields.length) {
    throw new Error('No writable business fields were found.');
  }
  return fields;
}

function getCopyValues(item: IListMigrationItem, fields: IListMigrationField[]): Record<string, unknown> {
  const values: Record<string, unknown> = {};
  for (const field of fields) {
    if (field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti') {
      continue;
    }
    if (item[field.InternalName] === undefined) {
      throw new Error(`Source value was not returned: ${field.InternalName}.`);
    }
    values[field.InternalName] = item[field.InternalName];
  }
  return values;
}

function normalizeValue(value: unknown): string {
  if (value === null || value === undefined) {
    return 'null';
  }
  if (Array.isArray(value)) {
    return JSON.stringify(value.map(normalizeValue).sort());
  }
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>;
    const normalized: Record<string, string> = {};
    for (const key of Object.keys(record).sort()) {
      normalized[key] = normalizeValue(record[key]);
    }
    return JSON.stringify(normalized);
  }
  return JSON.stringify(value);
}

export default class ListMigrationService {
  private source;
  private target;
  private userIds: Record<string, number> = {};
  private readonly userIdentityCache: Record<string, Promise<string>> = {};
  private running = false;
  private readonly request: IListMigrationRequest;

  public constructor(private context: WebPartContext, request: IListMigrationRequest) {
    this.request = {
      sourceSiteUrl: normalizeSiteUrl(request.sourceSiteUrl, context.pageContext.web.absoluteUrl),
      sourceListTitle: request.sourceListTitle.trim(),
      targetSiteUrl: normalizeSiteUrl(request.targetSiteUrl, context.pageContext.web.absoluteUrl),
      targetListTitle: request.targetListTitle.trim()
    };
    if (!this.request.sourceListTitle || !this.request.targetListTitle) {
      throw new Error('Enter both source and target list titles.');
    }
    this.source = spfi(this.request.sourceSiteUrl).using(SPFx(context));
    this.target = spfi(this.request.targetSiteUrl).using(SPFx(context));
  }

  private async checkPermissions(): Promise<void> {
    const list = this.target.web.lists.getByTitle(this.request.targetListTitle);
    for (const permission of [PermissionKind.AddListItems, PermissionKind.EditListItems]) {
      if (!await list.currentUserHasPermissions(permission)) {
        throw new Error('Add and edit item permissions are required on the target list.');
      }
    }
  }

  private getCheckpointKey(sourceId: string, targetId: string): string {
    const identity = [
      this.context.pageContext.user.loginName,
      this.request.sourceSiteUrl,
      sourceId,
      this.request.targetSiteUrl,
      targetId
    ].join('|');
    return `ListMigration.${encodeURIComponent(identity)}`;
  }

  public async preflight(): Promise<IListMigrationPlan> {
    const source = this.source.web.lists.getByTitle(this.request.sourceListTitle);
    const target = this.target.web.lists.getByTitle(this.request.targetListTitle);
    await this.checkPermissions();

    const sourceInfo = await source.select('Id', 'EnableAttachments')();
    const targetInfo = await target.select('Id', 'EnableAttachments')();
    if (this.request.sourceSiteUrl.toLowerCase() === this.request.targetSiteUrl.toLowerCase() &&
        sourceInfo.Id === targetInfo.Id) {
      throw new Error('Source and target must be different lists.');
    }
    if (sourceInfo.EnableAttachments && !targetInfo.EnableAttachments) {
      throw new Error('Enable attachments on the target list before migrating.');
    }

    const sourceFields = await source.fields() as IListMigrationField[];
    const targetFields = await target.fields() as IListMigrationField[];
    const fields = getCopyFields(sourceFields, targetFields);
    const selectedFields = fields.map(field =>
      field.TypeAsString.indexOf('User') === 0 ? `${field.InternalName}Id` : field.InternalName);
    const items: IListMigrationItem[] = [];
    for await (const page of source.items.select('Id', 'Attachments', ...selectedFields).orderBy('Id').top(500)) {
      items.push(...page as IListMigrationItem[]);
    }

    const targetItemIds = new Set<number>();
    for await (const page of target.items.select('Id').orderBy('Id').top(500)) {
      for (const item of page as Array<{ Id: number }>) {
        targetItemIds.add(item.Id);
      }
    }

    const key = this.getCheckpointKey(sourceInfo.Id, targetInfo.Id);
    const probeKey = `${key}.probe`;
    localStorage.setItem(probeKey, '1');
    localStorage.removeItem(probeKey);

    const checkpoint = this.readCheckpoint(key);
    let checkpointChanged = false;
    for (const sourceId of Object.keys(checkpoint)) {
      if (!targetItemIds.has(checkpoint[sourceId].targetId)) {
        delete checkpoint[sourceId];
        checkpointChanged = true;
      }
    }
    if (checkpointChanged) {
      if (Object.keys(checkpoint).length) {
        localStorage.setItem(key, JSON.stringify(checkpoint));
      } else {
        localStorage.removeItem(key);
      }
    }

    return {
      key,
      sourceListId: sourceInfo.Id,
      targetListId: targetInfo.Id,
      fields,
      items,
      targetCount: targetItemIds.size
    };
  }

  private readCheckpoint(key: string): IMigrationCheckpoint {
    const checkpoint: IMigrationCheckpoint = JSON.parse(localStorage.getItem(key) || '{}');
    if (!checkpoint || typeof checkpoint !== 'object' || Array.isArray(checkpoint) ||
        Object.keys(checkpoint).some(sourceId => !/^\d+$/.test(sourceId) ||
          !checkpoint[sourceId] || !Number.isInteger(checkpoint[sourceId].targetId) ||
          checkpoint[sourceId].targetId <= 0 || typeof checkpoint[sourceId].completed !== 'boolean')) {
      throw new Error('Invalid migration checkpoint. Reconcile target items before resetting it.');
    }
    return checkpoint;
  }

  private async getComparableValue(
    value: unknown,
    field: IListMigrationField,
    site: typeof this.source,
    siteKey: string
  ): Promise<string> {
    if (field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti') {
      const isMulti = field.TypeAsString === 'UserMulti';
      const ids = value === null || value === undefined ? [] : isMulti
        ? Array.isArray(value) ? value as number[] : [value as number]
        : [value as number];
      const identities: string[] = [];
      for (const id of ids) {
        if (!Number.isInteger(id) || id <= 0) {
          throw new Error(`Invalid person value while checking ${field.InternalName}.`);
        }
        const cacheKey = `${siteKey}:${id}`;
        if (!this.userIdentityCache[cacheKey]) {
          this.userIdentityCache[cacheKey] = site.web.siteUsers.getById(id)
            .select('Email', 'LoginName')()
            .then(person => (person.Email || person.LoginName).toLowerCase());
        }
        identities.push(await this.userIdentityCache[cacheKey]);
      }
      return JSON.stringify(isMulti ? identities.sort() : identities);
    }

    if (field.TypeAsString === 'DateTime' && value !== null && value !== undefined) {
      const date = new Date(value as string | number);
      if (!Number.isFinite(date.getTime())) {
        throw new Error(`Invalid date value while checking ${field.InternalName}.`);
      }
      return date.toISOString();
    }
    return normalizeValue(value);
  }

  public async verify(plan: IListMigrationPlan): Promise<IListVerificationReport> {
    const source = this.source.web.lists.getByTitle(this.request.sourceListTitle);
    const target = this.target.web.lists.getByTitle(this.request.targetListTitle);
    const [sourceInfo, targetInfo] = await Promise.all([
      source.select('Id')(),
      target.select('Id')()
    ]);
    if (sourceInfo.Id !== plan.sourceListId || targetInfo.Id !== plan.targetListId) {
      throw new Error('The verification plan no longer matches the selected lists. Run preflight again.');
    }

    const fieldsToSelect = plan.fields.map(field =>
      field.TypeAsString.indexOf('User') === 0 ? `${field.InternalName}Id` : field.InternalName);
    const sourceItems: IListMigrationItem[] = [];
    const targetItems: IListMigrationItem[] = [];
    for await (const page of source.items.select('Id', ...fieldsToSelect).orderBy('Id').top(500)) {
      sourceItems.push(...page as IListMigrationItem[]);
    }
    for await (const page of target.items.select('Id', ...fieldsToSelect).orderBy('Id').top(500)) {
      targetItems.push(...page as IListMigrationItem[]);
    }

    const getSignature = async (
      item: IListMigrationItem,
      site: typeof this.source,
      siteKey: string
    ): Promise<string> => {
      const values: Record<string, string> = {};
      for (const field of plan.fields) {
        const name = field.TypeAsString.indexOf('User') === 0
          ? `${field.InternalName}Id`
          : field.InternalName;
        if (item[name] === undefined) {
          throw new Error(`Could not read ${field.InternalName} from list item ${item.Id}.`);
        }
        values[field.InternalName] = await this.getComparableValue(item[name], field, site, siteKey);
      }
      return JSON.stringify(values);
    };

    const targetBuckets = new Map<string, IListMigrationItem[]>();
    for (const targetItem of targetItems) {
      const signature = await getSignature(targetItem, this.target, 'target');
      const bucket = targetBuckets.get(signature);
      if (bucket) {
        bucket.push(targetItem);
      } else {
        targetBuckets.set(signature, [targetItem]);
      }
    }

    const unmatchedSource: IListMigrationItem[] = [];
    const report: IListVerificationReport = {
      sourceRecords: sourceItems.length,
      targetRecords: targetItems.length,
      matchingRecords: 0,
      sourceOnlyRecords: 0,
      targetOnlyRecords: 0,
      differences: []
    };
    for (const sourceItem of sourceItems) {
      const signature = await getSignature(sourceItem, this.source, 'source');
      const bucket = targetBuckets.get(signature);
      if (bucket && bucket.length > 0) {
        bucket.shift();
        report.matchingRecords++;
      } else {
        unmatchedSource.push(sourceItem);
      }
    }

    const unmatchedTarget = Array.from(targetBuckets.values()).reduce(
      (items, bucket) => items.concat(bucket),
      [] as IListMigrationItem[]
    );
    report.sourceOnlyRecords = unmatchedSource.length;
    report.targetOnlyRecords = unmatchedTarget.length;
    for (const item of unmatchedSource) {
      report.differences.push({
        side: 'Source',
        itemId: item.Id,
        reason: 'No destination record with the same migrated field values.'
      });
    }
    for (const item of unmatchedTarget) {
      report.differences.push({
        side: 'Target',
        itemId: item.Id,
        reason: 'No source record with the same migrated field values.'
      });
    }
    return report;
  }

  public async migrate(plan: IListMigrationPlan, progress: (message: string) => void): Promise<void> {
    if (this.running) {
      throw new Error('A migration is already running.');
    }
    this.running = true;
    let activeId = 0;
    try {
      await this.checkPermissions();
      const source = this.source.web.lists.getByTitle(this.request.sourceListTitle);
      const target = this.target.web.lists.getByTitle(this.request.targetListTitle);
      const [sourceInfo, targetInfo] = await Promise.all([
        source.select('Id')(),
        target.select('Id')()
      ]);
      if (sourceInfo.Id !== plan.sourceListId || targetInfo.Id !== plan.targetListId ||
          plan.key !== this.getCheckpointKey(sourceInfo.Id, targetInfo.Id)) {
        throw new Error('The migration plan no longer matches the selected lists. Run preflight again.');
      }

      const checkpoint = this.readCheckpoint(plan.key);
      if (plan.targetCount > 0 && Object.keys(checkpoint).length === 0) {
        throw new Error('Target list is not empty and this browser has no migration checkpoint. Reconcile existing items before continuing.');
      }

      for (const item of plan.items) {
        activeId = item.Id;
        let entry = checkpoint[String(item.Id)];
        if (entry && entry.completed) {
          continue;
        }

        if (!entry) {
          const values = getCopyValues(item, plan.fields);
          for (const field of plan.fields.filter(candidate => candidate.TypeAsString.indexOf('User') === 0)) {
            const raw = item[`${field.InternalName}Id`];
            const multi = field.TypeAsString === 'UserMulti';
            const ids = raw === null ? [] : multi ? raw as number[] : [raw as number];
            if (!Array.isArray(ids) || ids.some(id => !Number.isInteger(id) || id <= 0)) {
              throw new Error(`Invalid person value: ${field.InternalName}.`);
            }
            const resolved: number[] = [];
            for (const userId of ids) {
              if (!this.userIds[String(userId)]) {
                const user = await this.source.web.siteUsers.getById(userId).select('Email', 'LoginName')();
                const targetUser = await this.target.web.ensureUser(user.Email || user.LoginName);
                this.userIds[String(userId)] = targetUser.Id;
              }
              resolved.push(this.userIds[String(userId)]);
            }
            values[`${field.InternalName}Id`] = multi ? resolved : resolved[0] || null;
          }

          const created = await target.items.add(values);
          if (!Number.isInteger(created.Id) || created.Id <= 0) {
            throw new Error('No target item ID was returned. Reconcile the target before retrying.');
          }
          entry = { targetId: created.Id, completed: false };
          checkpoint[String(item.Id)] = entry;
          localStorage.setItem(plan.key, JSON.stringify(checkpoint));
        }

        const targetItem = target.items.getById(entry.targetId);
        await targetItem.select('Id')();
        if (item.Attachments) {
          const existingAttachments = await targetItem.attachmentFiles();
          for (const attachment of await source.items.getById(item.Id).attachmentFiles()) {
            if (existingAttachments.some(candidate => candidate.FileName === attachment.FileName)) {
              continue;
            }
            const blob = await this.source.web
              .getFileByServerRelativePath(attachment.ServerRelativeUrl).getBlob();
            await targetItem.attachmentFiles.add(attachment.FileName, await blob.arrayBuffer());
          }
        }

        entry.completed = true;
        localStorage.setItem(plan.key, JSON.stringify(checkpoint));
        progress(`Source item ${item.Id} copied to target item ${entry.targetId}.`);
      }
      progress(`Migration completed for ${plan.items.length} source items. Verify target values and attachments.`);
    } catch (error) {
      throw new Error(`Migration stopped${activeId ? ` at source item ${activeId}` : ''}: ${
        error instanceof Error ? error.message : String(error)
      }`);
    } finally {
      this.running = false;
    }
  }
}
