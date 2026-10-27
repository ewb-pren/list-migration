# SharePoint List Migration

An SPFx web part for copying list item data and attachments between SharePoint sites in the same tenant.

## Requirements

- SharePoint Framework 1.22.2
- Access to read the source list and add/edit items in the destination list
- Existing source and destination lists with compatible fields

## Migration workflow

1. Enter the source and destination site URLs and list titles.
2. Run **Preflight** to validate list access, attachments, and field compatibility and to review the number of items.
3. Review the field mapping and target data, check the confirmation box, and start the migration.
4. Run **Check consistency** to compare all current source and destination records by the fields selected for migration.

The migration supports text, note, choice, multi-choice, number, currency, Boolean, date/time, URL, GUID, and person fields. It copies attachments when enabled on both lists. For person fields, source users are resolved in the destination site. SharePoint groups are mapped by group title where possible; unresolved group/principal values are omitted from the destination field and reported after migration rather than stopping other records. If the destination field is required and SharePoint rejects an item without that value, the item cannot be created until the group is mapped or the destination requirement is changed.

The web part counts current target list items during preflight rather than relying on SharePoint's potentially delayed list item count. Run preflight again after changing the target list; the displayed count is a snapshot from the last preflight. A browser-local checkpoint lets an interrupted run resume in the same browser profile. Preflight reconciles checkpoint entries whose target items have since been deleted, so those source items can be copied again. Do not clear local storage or manually remove target items while a migration is in progress. A non-empty target list without a checkpoint is intentionally blocked to reduce accidental duplication.

Consistency checking reads all current records in both lists and compares them as unordered collections using the participating migration field values. Duplicate rows are counted, and records are not matched by SharePoint item ID or browser-local checkpoints. Attachments and system fields are not part of the consistency comparison. If an allowed group/principal cannot be resolved, its value is excluded from comparison and a warning asks for a separate check; the remaining data is still compared. All records are held in browser memory during verification, so use caution with very large lists.

## Limitations

- The destination list must be created in advance. Its writable fields must have matching internal names and compatible types; required destination fields must be mapped or have defaults.
- Fields are matched by `InternalName`, while item queries, writes, and consistency checks use each site's `EntityPropertyName`. This supports encoded fields such as `_x0076_t09` whose REST property is `OData__x0076_t09`, including person fields with an `Id` suffix.
- This copies item values and attachments only. It does not copy list schema, views, content types, permissions, item IDs, version history, audit fields, or original link addresses.
- New destination items may trigger Power Automate flows, alerts, or other list automations.
- Review the destination and test a small migration before running a large or business-critical transfer.

## Build

```text
pnpm install
pnpm build
```

The web part uses PnPjs for SharePoint REST operations and Fluent UI for its interface.
