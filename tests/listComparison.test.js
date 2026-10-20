const { compareRecords } = require('../lib-commonjs/webparts/listMigration/services/ListComparison');

describe('compareRecords', () => {
  test('matches records independent of list ordering and item IDs', () => {
    const result = compareRecords(
      [
        { Id: 1, signature: '{"title":"Alpha"}' },
        { Id: 2, signature: '{"title":"Beta"}' }
      ],
      [
        { Id: 92, signature: '{"title":"Beta"}' },
        { Id: 91, signature: '{"title":"Alpha"}' }
      ]
    );

    expect(result).toEqual({
      matchingRecords: 2,
      sourceOnlyRecords: 0,
      targetOnlyRecords: 0,
      differences: []
    });
  });

  test('preserves duplicate counts and reports unmatched source and target rows', () => {
    const result = compareRecords(
      [
        { Id: 1, signature: 'duplicate' },
        { Id: 2, signature: 'duplicate' },
        { Id: 3, signature: 'source-only' }
      ],
      [
        { Id: 40, signature: 'duplicate' },
        { Id: 41, signature: 'target-only' }
      ]
    );

    expect(result).toEqual({
      matchingRecords: 1,
      sourceOnlyRecords: 2,
      targetOnlyRecords: 1,
      differences: [
        {
          side: 'Source',
          itemId: 2,
          reason: 'No destination record with the same migrated field values.'
        },
        {
          side: 'Source',
          itemId: 3,
          reason: 'No destination record with the same migrated field values.'
        },
        {
          side: 'Target',
          itemId: 41,
          reason: 'No source record with the same migrated field values.'
        }
      ]
    });
  });

  test('reports all records as unmatched when either list is empty', () => {
    expect(compareRecords([{ Id: 7, signature: 'source' }], [])).toMatchObject({
      matchingRecords: 0,
      sourceOnlyRecords: 1,
      targetOnlyRecords: 0
    });
    expect(compareRecords([], [{ Id: 8, signature: 'target' }])).toMatchObject({
      matchingRecords: 0,
      sourceOnlyRecords: 0,
      targetOnlyRecords: 1
    });
  });

  test('handles two empty lists as consistent', () => {
    expect(compareRecords([], [])).toEqual({
      matchingRecords: 0,
      sourceOnlyRecords: 0,
      targetOnlyRecords: 0,
      differences: []
    });
  });
});
