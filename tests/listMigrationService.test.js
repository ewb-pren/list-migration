jest.mock('@pnp/sp', () => ({ spfi: jest.fn(), SPFx: jest.fn() }));
jest.mock('@pnp/sp/security', () => ({
  PermissionKind: { AddListItems: 1, EditListItems: 2 }
}));
[
  '@pnp/sp/webs', '@pnp/sp/lists', '@pnp/sp/items', '@pnp/sp/fields',
  '@pnp/sp/attachments', '@pnp/sp/files', '@pnp/sp/site-users/web',
  '@pnp/sp/site-groups/web', '@pnp/sp/security/list'
].forEach(name => jest.doMock(name, () => ({})));

const { spfi } = require('@pnp/sp');
const ListMigrationService =
  require('../lib-commonjs/webparts/listMigration/services/ListMigrationService').default;

function field(InternalName, EntityPropertyName, TypeAsString = 'Text') {
  return {
    InternalName, EntityPropertyName, TypeAsString,
    Hidden: false, ReadOnlyField: false, SchemaXml: '<Field />'
  };
}

function list(Id, fields, records) {
  return {
    select: jest.fn(() => async () => ({ Id, EnableAttachments: false })),
    fields: jest.fn(async () => fields),
    currentUserHasPermissions: jest.fn(async () => true),
    items: {
      select: jest.fn(() => ({
        orderBy: () => ({
          top: () => ({
            async *[Symbol.asyncIterator]() { yield records; }
          })
        })
      })),
      add: jest.fn(async () => ({ Id: 91 })),
      getById: jest.fn(() => ({
        select: () => async () => ({ Id: 91 })
      }))
    }
  };
}

function setup(sourceFields, targetFields, sourceRecords, targetRecords = []) {
  const source = list('source-list', sourceFields, sourceRecords);
  const target = list('target-list', targetFields, targetRecords);
  const sourceWeb = { lists: { getByTitle: () => source } };
  const targetWeb = {
    lists: { getByTitle: () => target },
    ensureUser: jest.fn(async () => ({ Id: 42 }))
  };
  spfi.mockReturnValueOnce({ using: () => ({ web: sourceWeb }) })
    .mockReturnValueOnce({ using: () => ({ web: targetWeb }) });
  const service = new ListMigrationService({
    pageContext: {
      web: { absoluteUrl: 'https://example.sharepoint.com/sites/source' },
      user: { loginName: 'test-user' }
    }
  }, {
    sourceSiteUrl: 'https://example.sharepoint.com/sites/source',
    targetSiteUrl: 'https://example.sharepoint.com/sites/target',
    sourceListTitle: 'Configuration',
    targetListTitle: 'Configuration'
  });
  return { service, source, target, targetWeb };
}

beforeEach(() => {
  jest.clearAllMocks();
  localStorage.clear();
});

test('preflight selects the encoded Body REST property and retains both property names', async () => {
  const { service, source } = setup(
    [field('_x0076_t09', 'OData__x0076_t09', 'Note')],
    [field('_x0076_t09', 'TargetBody', 'Note')],
    [{ Id: 1, Attachments: false, OData__x0076_t09: '<p>Body</p>' }]
  );
  const plan = await service.preflight();

  expect(source.items.select).toHaveBeenCalledWith('Id', 'Attachments', 'OData__x0076_t09');
  expect(plan.fields[0]).toMatchObject({
    InternalName: '_x0076_t09',
    EntityPropertyName: 'OData__x0076_t09',
    TargetEntityPropertyName: 'TargetBody'
  });
  await service.migrate(plan, jest.fn());
});

test('migration reads source REST properties and writes target properties without losing nulls or falsy values', async () => {
  const sourceFields = [
    field('_x0076_t09', 'OData__x0076_t09', 'Note'),
    field('Title', 'Title'),
    field('Enabled', 'Enabled', 'Boolean'),
    field('Amount', 'Amount', 'Number'),
    field('Optional', 'Optional')
  ];
  const targetFields = sourceFields.map(value => ({
    ...value,
    EntityPropertyName: value.InternalName === '_x0076_t09' ? 'TargetBody' : value.EntityPropertyName
  }));
  const { service, target } = setup(sourceFields, targetFields, [{
    Id: 1, Attachments: false, OData__x0076_t09: '<p>Body</p>',
    Title: '', Enabled: false, Amount: 0, Optional: null
  }]);

  await service.migrate(await service.preflight(), jest.fn());

  expect(target.items.add).toHaveBeenCalledWith({
    TargetBody: '<p>Body</p>', Title: '', Enabled: false, Amount: 0, Optional: null
  });
});

test.each(['User', 'UserMulti'])('maps encoded %s Id properties on reads, writes and verification', async type => {
  const multi = type === 'UserMulti';
  const { service, source, target, targetWeb } = setup(
    [field('_Owner', 'OData__Owner', type)],
    [field('_Owner', 'TargetOwner', type)],
    [{ Id: 1, Attachments: false, OData__OwnerId: multi ? [7] : 7 }]
  );
  jest.spyOn(service, 'resolvePrincipal').mockResolvedValue({
    kind: 'user', identity: 'user@example.com', title: 'User'
  });
  const plan = await service.preflight();
  expect(source.items.select).toHaveBeenCalledWith('Id', 'Attachments', 'OData__OwnerId');
  await service.migrate(plan, jest.fn());
  expect(targetWeb.ensureUser).toHaveBeenCalledWith('user@example.com');
  expect(target.items.add).toHaveBeenCalledWith({ TargetOwnerId: multi ? [42] : 42 });

  const targetRecords = [{ Id: 91, TargetOwnerId: multi ? [42] : 42 }];
  target.items.select.mockImplementation(list('target-list', [], targetRecords).items.select);
  expect(await service.verify(plan)).toMatchObject({ matchingRecords: 1, differences: [] });
  expect(target.items.select).toHaveBeenLastCalledWith('Id', 'TargetOwnerId');
});

test('verification selects each side independently and compares by internal field identity', async () => {
  const { service, source, target } = setup(
    [field('_x0076_t09', 'OData__x0076_t09', 'Note')],
    [field('_x0076_t09', 'TargetBody', 'Note')],
    [{ Id: 1, OData__x0076_t09: 'same body' }],
    [{ Id: 91, TargetBody: 'same body' }]
  );
  const report = await service.verify(await service.preflight());

  expect(source.items.select).toHaveBeenLastCalledWith('Id', 'OData__x0076_t09');
  expect(target.items.select).toHaveBeenLastCalledWith('Id', 'TargetBody');
  expect(report).toMatchObject({
    matchingRecords: 1, sourceOnlyRecords: 0, targetOnlyRecords: 0, differences: []
  });
});

test.each(['source', 'target'])('preflight rejects missing %s REST metadata instead of guessing', async side => {
  const { service, source } = setup(
    [field('Body', side === 'source' ? undefined : 'SourceBody')],
    [field('Body', side === 'target' ? undefined : 'TargetBody')],
    []
  );
  await expect(service.preflight()).rejects.toThrow('Missing REST property name for source or target field: Body.');
  expect(source.items.select).not.toHaveBeenCalled();
});

test('migration explicitly rejects an absent source value without creating a destination item', async () => {
  const { service, target } = setup(
    [field('_x0076_t09', 'OData__x0076_t09')],
    [field('_x0076_t09', 'TargetBody')],
    [{ Id: 1, Attachments: false }]
  );
  await expect(service.migrate(await service.preflight(), jest.fn()))
    .rejects.toThrow('Source value was not returned: _x0076_t09.');
  expect(target.items.add).not.toHaveBeenCalled();
});
