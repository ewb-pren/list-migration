import { spfi, SPFx } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PermissionKind } from '@pnp/sp/security';
import { compareRecords } from './ListComparison';
import { getReadableErrorMessage } from './errorHandling';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/attachments';
import '@pnp/sp/files';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/security/list';

export interface IListMigrationRequest {
  sourceSiteUrl: string;
  sourceListTitle: string;
  targetSiteUrl: string;
  targetListTitle: string;
}

export interface IListMigrationField {
  InternalName: string;
  EntityPropertyName: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
  Required?: boolean;
  DefaultValue?: string;
  SchemaXml: string;
}

export interface IListMigrationMappedField extends IListMigrationField {
  TargetEntityPropertyName: string;
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
  fields: IListMigrationMappedField[];
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
  groupWarnings: string[];
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

function getCopyFields(source: IListMigrationField[], target: IListMigrationField[]): IListMigrationMappedField[] {
  const fields = source.filter(field => !field.Hidden && !field.ReadOnlyField &&
    excludedFields.indexOf(field.InternalName) < 0);
  const mappedFields: IListMigrationMappedField[] = [];

  for (const field of fields) {
    const destination = target.find(candidate => candidate.InternalName === field.InternalName);
    if (supportedTypes.indexOf(field.TypeAsString) < 0) {
      throw new Error(`Unsupported field: ${field.InternalName} (${field.TypeAsString}).`);
    }
    if (!destination || destination.ReadOnlyField || destination.TypeAsString !== field.TypeAsString) {
      throw new Error(`Missing or incompatible target field: ${field.InternalName}.`);
    }
    if (!field.EntityPropertyName || !destination.EntityPropertyName) {
      throw new Error(`Missing REST property name for source or target field: ${field.InternalName}.`);
    }
    mappedFields.push({ ...field, TargetEntityPropertyName: destination.EntityPropertyName });
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
  return mappedFields;
}

function getItemPropertyName(field: IListMigrationMappedField, side: 'source' | 'target'): string {
  const name = side === 'source' ? field.EntityPropertyName : field.TargetEntityPropertyName;
  return field.TypeAsString.indexOf('User') === 0 ? `${name}Id` : name;
}

function getCopyValues(item: IListMigrationItem, fields: IListMigrationMappedField[]): Record<string, unknown> {
  const values: Record<string, unknown> = {};
  for (const field of fields) {
    if (field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti') {
      continue;
    }
    if (item[field.EntityPropertyName] === undefined) {
      throw new Error(`Source value was not returned: ${field.InternalName}.`);
    }
    values[field.TargetEntityPropertyName] = item[field.EntityPropertyName];
  }
  return values;
}

function fieldAllowsGroups(field: IListMigrationField): boolean {
  const schema = new DOMParser().parseFromString(field.SchemaXml, 'application/xml');
  if (schema.querySelector('parsererror')) {
    return false;
  }
  const mode = (schema.documentElement.getAttribute('UserSelectionMode') || '').toLowerCase();
  return mode === 'peopleandgroups' || mode === '1';
}

function isNotFoundError(error: unknown): boolean {
  const message = getReadableErrorMessage(error);
  return /HTTP 404\b/i.test(message) || /cannot be found/i.test(message);
}

function isGroupPrincipal(principalType: unknown): boolean {
  return typeof principalType === 'number' && principalType > 0 && (principalType & 1) === 0;
}

interface IResolvedPrincipal {
  kind: 'user' | 'group';
  identity: string;
  title: string;
}

class PrincipalNotFoundError extends Error {}

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
  private readonly principalCache: Record<string, Promise<IResolvedPrincipal>> = {};
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
    try {
      return await this.runPreflight();
    } catch (error) {
      throw new Error(
        `Could not preflight source list "${this.request.sourceListTitle}" at ${this.request.sourceSiteUrl} ` +
        `and destination list "${this.request.targetListTitle}" at ${this.request.targetSiteUrl}. ` +
        getReadableErrorMessage(error)
      );
    }
  }

  private async runPreflight(): Promise<IListMigrationPlan> {
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
    const selectedFields = fields.map(field => getItemPropertyName(field, 'source'));
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
    this.checkLocalStorage(probeKey);

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
        this.saveCheckpoint(key, checkpoint);
      } else {
        try {
          localStorage.removeItem(key);
        } catch (error) {
          throw new Error(
            'Could not update the browser migration checkpoint after reconciling deleted target records. ' +
            getReadableErrorMessage(error)
          );
        }
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
    let checkpoint: IMigrationCheckpoint;
    try {
      checkpoint = JSON.parse(localStorage.getItem(key) || '{}') as IMigrationCheckpoint;
    } catch (error) {
      throw new Error(
        `The saved migration checkpoint is unreadable. Do not delete it before reconciling destination records. ` +
        getReadableErrorMessage(error)
      );
    }
    if (!checkpoint || typeof checkpoint !== 'object' || Array.isArray(checkpoint) ||
        Object.keys(checkpoint).some(sourceId => !/^\d+$/.test(sourceId) ||
          !checkpoint[sourceId] || !Number.isInteger(checkpoint[sourceId].targetId) ||
          checkpoint[sourceId].targetId <= 0 || typeof checkpoint[sourceId].completed !== 'boolean')) {
      throw new Error(
        'The saved migration checkpoint has an invalid format. Reconcile destination records before resetting it.'
      );
    }
    return checkpoint;
  }

  private checkLocalStorage(probeKey: string): void {
    try {
      localStorage.setItem(probeKey, '1');
      localStorage.removeItem(probeKey);
    } catch (error) {
      throw new Error(
        'This browser cannot access local storage, which is required for migration recovery. ' +
        'Enable browser storage for this SharePoint site and retry. ' +
        getReadableErrorMessage(error)
      );
    }
  }

  private saveCheckpoint(key: string, checkpoint: IMigrationCheckpoint): void {
    try {
      localStorage.setItem(key, JSON.stringify(checkpoint));
    } catch (error) {
      throw new Error(
        'Could not save migration progress in this browser. Check that local storage is enabled and has free space. ' +
        'The destination may already contain the current item; reconcile it before retrying. ' +
        getReadableErrorMessage(error)
      );
    }
  }

  private resolvePrincipal(site: typeof this.source, siteKey: string, id: number): Promise<IResolvedPrincipal> {
    const cacheKey = `${siteKey}:${id}`;
    if (!this.principalCache[cacheKey]) {
      this.principalCache[cacheKey] = (async (): Promise<IResolvedPrincipal> => {
        try {
          const user = await site.web.siteUsers.getById(id).select('Email', 'LoginName', 'Title', 'PrincipalType')();
          const isGroup = isGroupPrincipal(user.PrincipalType);
          const identity = isGroup
            ? user.Title || user.LoginName || user.Email
            : user.Email || user.LoginName || user.Title;
          if (!identity) {
            throw new Error(`SharePoint principal ID ${id} has no usable identity.`);
          }
          return { kind: isGroup ? 'group' : 'user', identity, title: user.Title || identity };
        } catch (userError) {
          if (!isNotFoundError(userError)) {
            throw userError;
          }
        }
        // SharePoint groups are not returned by siteUsers, so fall back to siteGroups.
        try {
          const group = await site.web.siteGroups.getById(id).select('Title', 'LoginName')();
          const title = group.Title || group.LoginName;
          if (!title) {
            throw new Error(`SharePoint group ID ${id} has no title.`);
          }
          return { kind: 'group', identity: title, title };
        } catch (groupError) {
          if (isNotFoundError(groupError)) {
            throw new PrincipalNotFoundError(
              `No user or SharePoint group with ID ${id} exists in the ${siteKey} site. ` +
              getReadableErrorMessage(groupError)
            );
          }
          throw groupError;
        }
      })();
    }
    return this.principalCache[cacheKey];
  }

  private async getComparableValue(
    value: unknown,
    field: IListMigrationField,
    site: typeof this.source,
    siteKey: string,
    itemId: number,
    groupWarnings: Set<string>
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
        try {
          const principal = await this.resolvePrincipal(site, siteKey, id);
          identities.push(principal.kind === 'group'
            ? `group:${principal.title.toLowerCase()}`
            : `user:${principal.identity.toLowerCase()}`);
        } catch (error) {
          if (error instanceof PrincipalNotFoundError && fieldAllowsGroups(field)) {
            groupWarnings.add(
              `Could not resolve user/group ID ${id} in ${siteKey} item ${itemId}, field "${field.InternalName}"; ` +
              'this value was excluded from comparison and needs a separate check.'
            );
            continue;
          }
          throw new Error(
            `Unable to resolve user/group ID ${id} in the ${siteKey} site for field "${field.InternalName}". ` +
            'Confirm the account or group exists and is available in that SharePoint site. ' +
            `SharePoint detail: ${getReadableErrorMessage(error)}`
          );
        }
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
    try {
      return await this.runVerification(plan);
    } catch (error) {
      throw new Error(
        `Could not compare source list "${this.request.sourceListTitle}" at ${this.request.sourceSiteUrl} ` +
        `with destination list "${this.request.targetListTitle}" at ${this.request.targetSiteUrl}. ` +
        getReadableErrorMessage(error)
      );
    }
  }

  private async runVerification(plan: IListMigrationPlan): Promise<IListVerificationReport> {
    const source = this.source.web.lists.getByTitle(this.request.sourceListTitle);
    const target = this.target.web.lists.getByTitle(this.request.targetListTitle);
    const [sourceInfo, targetInfo] = await Promise.all([
      source.select('Id')(),
      target.select('Id')()
    ]);
    if (sourceInfo.Id !== plan.sourceListId || targetInfo.Id !== plan.targetListId) {
      throw new Error('The verification plan no longer matches the selected lists. Run preflight again.');
    }

    const sourceFieldsToSelect = plan.fields.map(field => getItemPropertyName(field, 'source'));
    const targetFieldsToSelect = plan.fields.map(field => getItemPropertyName(field, 'target'));
    const sourceItems: IListMigrationItem[] = [];
    const targetItems: IListMigrationItem[] = [];
    for await (const page of source.items.select('Id', ...sourceFieldsToSelect).orderBy('Id').top(500)) {
      sourceItems.push(...page as IListMigrationItem[]);
    }
    for await (const page of target.items.select('Id', ...targetFieldsToSelect).orderBy('Id').top(500)) {
      targetItems.push(...page as IListMigrationItem[]);
    }

    const getSignature = async (
      item: IListMigrationItem,
      site: typeof this.source,
      siteKey: 'source' | 'target',
      groupWarnings: Set<string>
    ): Promise<string> => {
      const values: Record<string, string> = {};
      for (const field of plan.fields) {
        const name = getItemPropertyName(field, siteKey);
        if (item[name] === undefined) {
          throw new Error(
            `Field "${field.InternalName}" was not returned for item ID ${item.Id} ` +
            `from the ${siteKey} list "${siteKey === 'source' ? this.request.sourceListTitle : this.request.targetListTitle}".`
          );
        }
        try {
          values[field.InternalName] = await this.getComparableValue(
            item[name], field, site, siteKey, item.Id, groupWarnings
          );
        } catch (error) {
          throw new Error(
            `Could not compare field "${field.InternalName}" for item ID ${item.Id} ` +
            `in the ${siteKey} list "${siteKey === 'source' ? this.request.sourceListTitle : this.request.targetListTitle}". ` +
            getReadableErrorMessage(error)
          );
        }
      }
      return JSON.stringify(values);
    };

    const groupWarnings = new Set<string>();
    const [sourceRecords, targetRecords] = await Promise.all([
      Promise.all(sourceItems.map(async item => ({
        Id: item.Id,
        signature: await getSignature(item, this.source, 'source', groupWarnings)
      }))),
      Promise.all(targetItems.map(async item => ({
        Id: item.Id,
        signature: await getSignature(item, this.target, 'target', groupWarnings)
      })))
    ]);
    return {
      sourceRecords: sourceRecords.length,
      targetRecords: targetRecords.length,
      ...compareRecords(sourceRecords, targetRecords),
      groupWarnings: Array.from(groupWarnings)
    };
  }

  public async migrate(plan: IListMigrationPlan, progress: (message: string) => void): Promise<string[]> {
    if (this.running) {
      throw new Error('A migration is already running.');
    }
    this.running = true;
    let activeId = 0;
    const groupWarnings = new Set<string>();
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
            const raw = item[getItemPropertyName(field, 'source')];
            const multi = field.TypeAsString === 'UserMulti';
            const ids = raw === null ? [] : multi ? raw as number[] : [raw as number];
            if (!Array.isArray(ids) || ids.some(id => !Number.isInteger(id) || id <= 0)) {
              throw new Error(`Invalid person value: ${field.InternalName}.`);
            }
            const resolved: number[] = [];
            for (const userId of ids) {
              if (!this.userIds[String(userId)]) {
                let principal: IResolvedPrincipal;
                try {
                  principal = await this.resolvePrincipal(this.source, 'source', userId);
                } catch (error) {
                  if (error instanceof PrincipalNotFoundError && fieldAllowsGroups(field)) {
                    groupWarnings.add(
                      `Could not resolve user/group ID ${userId} for source item ${item.Id}, field ` +
                      `"${field.InternalName}"; the value was left empty and needs a separate check.`
                    );
                    continue;
                  }
                  throw new Error(
                    `Could not map person field "${field.InternalName}" for source item ${item.Id} ` +
                    `(source user/group ID ${userId}). Confirm the source account exists and the destination site can resolve it. ` +
                    getReadableErrorMessage(error)
                  );
                }
                try {
                  const targetPrincipal = principal.kind === 'group'
                    ? await this.target.web.siteGroups.getByName(principal.title).select('Id')()
                    : await this.target.web.ensureUser(principal.identity);
                  this.userIds[String(userId)] = targetPrincipal.Id;
                } catch (error) {
                  if (principal.kind === 'group') {
                    groupWarnings.add(
                      `Could not resolve group "${principal.title}" in the destination site for source item ${item.Id}, ` +
                      `field "${field.InternalName}" (source ID ${userId}); the value was left empty and ` +
                      'needs a separate check.'
                    );
                    continue;
                  }
                  throw new Error(
                    `Could not map person field "${field.InternalName}" for source item ${item.Id} ` +
                    `(source user ID ${userId}, "${principal.identity}"). The destination site could not resolve this account. ` +
                    getReadableErrorMessage(error)
                  );
                }
              }
              resolved.push(this.userIds[String(userId)]);
            }
            if (resolved.length || !fieldAllowsGroups(field)) {
              values[getItemPropertyName(field, 'target')] = multi ? resolved : resolved[0] || null;
            }
          }

          const created = await target.items.add(values);
          if (!Number.isInteger(created.Id) || created.Id <= 0) {
            throw new Error('No target item ID was returned. Reconcile the target before retrying.');
          }
          entry = { targetId: created.Id, completed: false };
          checkpoint[String(item.Id)] = entry;
          this.saveCheckpoint(plan.key, checkpoint);
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
        this.saveCheckpoint(plan.key, checkpoint);
        progress(`Source item ${item.Id} copied to target item ${entry.targetId}.`);
      }
      progress(`Migration completed for ${plan.items.length} source items. Verify target values and attachments.`);
      return Array.from(groupWarnings);
    } catch (error) {
      throw new Error(
        `Migration stopped${activeId ? ` at source item ${activeId}` : ''} for destination list ` +
        `"${this.request.targetListTitle}" at ${this.request.targetSiteUrl}: ${
        getReadableErrorMessage(error)
        }`
      );
    } finally {
      this.running = false;
    }
  }
}
