const { getReadableErrorMessage } = require('../lib-commonjs/webparts/listMigration/services/errorHandling');

describe('getReadableErrorMessage', () => {
  test('extracts the SharePoint OData message from a PnPjs HTTP error', () => {
    const error = new Error(
      'Error making HttpClient request in queryable [404] ::> ' +
      '{"odata.error":{"code":"-1, Microsoft.SharePoint.Client.ResourceNotFoundException",' +
      '"message":{"lang":"en-US","value":"User cannot be found."}}}'
    );

    expect(getReadableErrorMessage(error)).toBe('HTTP 404: User cannot be found.');
  });

  test('includes HTTP status for errors with a response object', () => {
    expect(getReadableErrorMessage({
      message: 'Request failed',
      response: { status: 403, statusText: 'Forbidden' }
    })).toBe('HTTP 403 Forbidden: Request failed');
  });

  test('extracts server messages from structured response bodies', () => {
    expect(getReadableErrorMessage({
      message: 'Request failed',
      response: {
        status: 404,
        data: {
          'odata.error': {
            message: { value: 'The list item was not found.' }
          }
        }
      }
    })).toBe('HTTP 404: The list item was not found.');
  });

  test('preserves ordinary error messages', () => {
    expect(getReadableErrorMessage(new Error('The destination list is missing.')))
      .toBe('The destination list is missing.');
  });
});
