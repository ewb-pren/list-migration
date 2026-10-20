export interface IComparableRecord {
  Id: number;
  signature: string;
}

export interface IRecordDifference {
  side: 'Source' | 'Target';
  itemId: number;
  reason: string;
}

export interface IRecordComparison {
  matchingRecords: number;
  sourceOnlyRecords: number;
  targetOnlyRecords: number;
  differences: IRecordDifference[];
}

export function compareRecords(
  sourceRecords: IComparableRecord[],
  targetRecords: IComparableRecord[]
): IRecordComparison {
  const targetBuckets = new Map<string, IComparableRecord[]>();
  for (const targetRecord of targetRecords) {
    const bucket = targetBuckets.get(targetRecord.signature);
    if (bucket) {
      bucket.push(targetRecord);
    } else {
      targetBuckets.set(targetRecord.signature, [targetRecord]);
    }
  }

  let matchingRecords = 0;
  const sourceOnly: IComparableRecord[] = [];
  for (const sourceRecord of sourceRecords) {
    const bucket = targetBuckets.get(sourceRecord.signature);
    if (bucket && bucket.length > 0) {
      bucket.shift();
      matchingRecords++;
    } else {
      sourceOnly.push(sourceRecord);
    }
  }

  const targetOnly = Array.from(targetBuckets.values()).reduce(
    (records, bucket) => records.concat(bucket),
    [] as IComparableRecord[]
  );
  return {
    matchingRecords,
    sourceOnlyRecords: sourceOnly.length,
    targetOnlyRecords: targetOnly.length,
    differences: [
      ...sourceOnly.map(record => ({
        side: 'Source' as const,
        itemId: record.Id,
        reason: 'No destination record with the same migrated field values.'
      })),
      ...targetOnly.map(record => ({
        side: 'Target' as const,
        itemId: record.Id,
        reason: 'No source record with the same migrated field values.'
      }))
    ]
  };
}
