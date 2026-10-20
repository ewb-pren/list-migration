interface IErrorRecord {
  message?: unknown;
  status?: unknown;
  statusCode?: unknown;
  response?: {
    status?: unknown;
    statusText?: unknown;
    data?: unknown;
    text?: unknown;
  };
  data?: unknown;
}

function getString(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim() ? value.trim() : undefined;
}

function parseSharePointMessage(value: string): string | undefined {
  const jsonStart = value.indexOf('{');
  if (jsonStart < 0) {
    return undefined;
  }

  try {
    const parsed = JSON.parse(value.slice(jsonStart)) as {
      'odata.error'?: { message?: { value?: string } };
      error?: { message?: string };
      message?: string;
    };
    return getString(parsed['odata.error']?.message?.value) ||
      getString(parsed.error?.message) ||
      getString(parsed.message);
  } catch {
    return undefined;
  }
}

function getServerMessage(value: unknown): string | undefined {
  const text = getString(value);
  if (text) {
    return parseSharePointMessage(text);
  }
  if (!value || typeof value !== 'object') {
    return undefined;
  }

  const record = value as {
    'odata.error'?: { message?: { value?: string } | string };
    error?: { message?: string } | string;
    message?: { value?: string } | string;
    data?: unknown;
    text?: unknown;
  };
  const odataMessage = record['odata.error']?.message;
  const errorMessage = record.error && typeof record.error === 'object'
    ? record.error.message
    : record.error;
  const message = record.message && typeof record.message === 'object'
    ? record.message.value
    : record.message;
  const directMessage = typeof odataMessage === 'object'
    ? getString(odataMessage?.value)
    : getString(odataMessage);
  return directMessage || getString(errorMessage) || getString(message) ||
    getServerMessage(record.data) || getServerMessage(record.text);
}

export function getReadableErrorMessage(error: unknown): string {
  if (typeof error === 'string') {
    return getServerMessage(error) || error;
  }
  if (!error || typeof error !== 'object') {
    return String(error);
  }

  const record = error as IErrorRecord;
  const message = getString(record.message) || String(error);
  const serverMessage = getServerMessage(message) ||
    getServerMessage(record.response?.text) ||
    getServerMessage(record.response?.data) ||
    getServerMessage(record.data);
  const messageStatus = /\[(\d{3})\]/.exec(message)?.[1];
  const status = record.response?.status ?? record.status ?? record.statusCode ?? messageStatus;
  const statusText = getString(record.response?.statusText);
  const statusLabel = status ? `HTTP ${String(status)}${statusText ? ` ${statusText}` : ''}` : '';

  if (serverMessage) {
    return statusLabel ? `${statusLabel}: ${serverMessage}` : serverMessage;
  }

  const cleanMessage = message.replace(/\s*::>\s*\{[\s\S]*$/, '').trim();
  return statusLabel && !cleanMessage.includes(statusLabel)
    ? `${statusLabel}: ${cleanMessage}`
    : cleanMessage;
}
