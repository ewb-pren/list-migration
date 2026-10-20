import * as React from 'react';
import {
  Checkbox,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  Stack,
  TextField
} from '@fluentui/react';
import styles from './ListMigration.module.scss';
import type { IListMigrationProps } from './IListMigrationProps';
import ListMigrationService, {
  IListVerificationReport,
  IListMigrationPlan,
  IListMigrationRequest
} from '../services/ListMigrationService';
import { getReadableErrorMessage } from '../services/errorHandling';

export default function ListMigration(props: IListMigrationProps): React.ReactElement {
  const currentSiteUrl = props.context.pageContext.web.absoluteUrl;
  const [sourceSiteUrl, setSourceSiteUrl] = React.useState(currentSiteUrl);
  const [targetSiteUrl, setTargetSiteUrl] = React.useState(currentSiteUrl);
  const [sourceListTitle, setSourceListTitle] = React.useState('');
  const [targetListTitle, setTargetListTitle] = React.useState('');
  const [plan, setPlan] = React.useState<IListMigrationPlan>();
  const [busy, setBusy] = React.useState(false);
  const [confirmed, setConfirmed] = React.useState(false);
  const [migrationSucceeded, setMigrationSucceeded] = React.useState(false);
  const [migrationGroupWarnings, setMigrationGroupWarnings] = React.useState<string[]>([]);
  const [verificationReport, setVerificationReport] = React.useState<IListVerificationReport>();
  const [error, setError] = React.useState('');
  const [messages, setMessages] = React.useState<string[]>([]);
  const [pendingOperation, setPendingOperation] = React.useState<'preflight' | 'migrate' | 'verify'>();
  const inFlight = React.useRef(false);
  const releaseInFlight = (): void => {
    inFlight.current = false;
  };

  const resetPlan = (): void => {
    setPlan(undefined);
    setConfirmed(false);
    setMigrationSucceeded(false);
    setMigrationGroupWarnings([]);
    setVerificationReport(undefined);
    setMessages([]);
  };

  const updateRequestValue = (setter: (value: string) => void, value?: string): void => {
    setter(value || '');
    resetPlan();
    setError('');
  };

  const getRequest = (): IListMigrationRequest => ({
    sourceSiteUrl,
    sourceListTitle,
    targetSiteUrl,
    targetListTitle
  });

  const isProductionRisk = !/(dev|uat)/i.test(targetSiteUrl);

  const runPreflight = async (): Promise<void> => {
    if (inFlight.current) return;
    inFlight.current = true;
    setBusy(true);
    resetPlan();
    setError('');
    try {
      const service = new ListMigrationService(props.context, getRequest());
      setPlan(await service.preflight());
    } catch (failure) {
      setError(`Preflight failed:\n${getReadableErrorMessage(failure)}`);
    } finally {
      releaseInFlight();
      setBusy(false);
    }
  };

  const runMigration = async (): Promise<void> => {
    if (!plan || !confirmed || inFlight.current) return;
    inFlight.current = true;
    setBusy(true);
    setMigrationSucceeded(false);
    setVerificationReport(undefined);
    setError('');
    try {
      const service = new ListMigrationService(props.context, getRequest());
      setMigrationGroupWarnings(
        await service.migrate(plan, message => setMessages(previous => [...previous.slice(-29), message]))
      );
      setConfirmed(false);
      setMigrationSucceeded(true);
    } catch (failure) {
      setError(`Migration failed:\n${getReadableErrorMessage(failure)}`);
    } finally {
      releaseInFlight();
      setBusy(false);
    }
  };

  const runVerification = async (): Promise<void> => {
    if (!plan || inFlight.current) return;
    inFlight.current = true;
    setBusy(true);
    setVerificationReport(undefined);
    setError('');
    try {
      const service = new ListMigrationService(props.context, getRequest());
      setVerificationReport(await service.verify(plan));
    } catch (failure) {
      setError(`Consistency check failed:\n${getReadableErrorMessage(failure)}`);
    } finally {
      releaseInFlight();
      setBusy(false);
    }
  };

  const requestOperation = async (operation: 'preflight' | 'migrate' | 'verify'): Promise<void> => {
    if (isProductionRisk) {
      setPendingOperation(operation);
      return;
    }
    if (operation === 'preflight') {
      await runPreflight();
    } else if (operation === 'migrate') {
      await runMigration();
    } else {
      await runVerification();
    }
  };

  const confirmOperation = async (): Promise<void> => {
    const operation = pendingOperation;
    setPendingOperation(undefined);
    if (operation === 'preflight') {
      await runPreflight();
    } else if (operation === 'migrate') {
      await runMigration();
    } else if (operation === 'verify') {
      await runVerification();
    }
  };

  const cellStyle: React.CSSProperties = {
    padding: '8px 12px',
    borderBottom: '1px solid #edebe9',
    textAlign: 'left',
    overflowWrap: 'anywhere'
  };

  return (
    <section className={styles.listMigration} aria-label="SharePoint list migration">
      <header className={styles.header}>
        <div>
          <p className={styles.eyebrow}>SHAREPOINT DATA TRANSFER</p>
          <h1>List migration</h1>
          <p className={styles.subtitle}>Copy list items and attachments between sites in your tenant.</p>
        </div>
        <span className={styles.headerIcon} aria-hidden="true">⇄</span>
      </header>

      <div className={styles.card}>
        <h2>Migration source and destination</h2>
        <div className={styles.siteGrid}>
          <div className={styles.sitePanel}>
            <div className={styles.stepLabel}><span>1</span> SOURCE</div>
            <TextField
              label="Source site URL"
              value={sourceSiteUrl}
              disabled={busy}
              onChange={(_event, value) => updateRequestValue(setSourceSiteUrl, value)}
              placeholder="https://tenant.sharepoint.com/sites/source"
              description="Use an HTTPS site URL in this SharePoint tenant."
            />
            <TextField
              label="Source list title"
              value={sourceListTitle}
              disabled={busy}
              onChange={(_event, value) => updateRequestValue(setSourceListTitle, value)}
              placeholder="e.g. Requests"
            />
          </div>
          <div className={styles.sitePanel}>
            <div className={styles.stepLabel}><span>2</span> DESTINATION</div>
            <TextField
              label="Target site URL"
              value={targetSiteUrl}
              disabled={busy}
              onChange={(_event, value) => updateRequestValue(setTargetSiteUrl, value)}
              placeholder="https://tenant.sharepoint.com/sites/target"
              description="The destination list must already exist."
            />
            <TextField
              label="Target list title"
              value={targetListTitle}
              disabled={busy}
              onChange={(_event, value) => updateRequestValue(setTargetListTitle, value)}
              placeholder="e.g. Requests"
            />
          </div>
        </div>

        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          This migrates item values and attachments into an existing compatible list. It does not create the list or
          preserve item IDs, version history, permissions, audit fields, content types, or embedded links. Target
          workflows may run when items are added.
        </MessageBar>

        <Stack horizontal wrap tokens={{ childrenGap: 12 }} className={styles.actions}>
          <DefaultButton
            text="Run preflight"
            iconProps={{ iconName: 'CheckList' }}
            disabled={busy}
            onClick={() => requestOperation('preflight')}
          />
          <PrimaryButton
            text="Start migration"
            iconProps={{ iconName: 'Copy' }}
            disabled={busy || !plan || !confirmed || !plan.items.length}
            onClick={() => requestOperation('migrate')}
          />
          <DefaultButton
            text="Check consistency"
            iconProps={{ iconName: 'Compare' }}
            disabled={busy || !plan}
            onClick={() => requestOperation('verify')}
          />
          {busy && <Spinner label="Processing migration" />}
        </Stack>

        {error && (
          <MessageBar messageBarType={MessageBarType.error} role="alert" isMultiline>
            <div style={{ whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>{error}</div>
          </MessageBar>
        )}

        {plan && <>
          <div className={styles.summary}>
            <div><span>Source items</span><strong>{plan.items.length}</strong></div>
            <div><span>Existing target items (preflight count)</span><strong>{plan.targetCount}</strong></div>
            <div><span>Mapped fields</span><strong>{plan.fields.length}</strong></div>
          </div>
          <h3>Field compatibility</h3>
          <div className={styles.tableWrap}>
            <table className={styles.fieldTable} aria-label="Compatible fields">
              <thead><tr><th style={cellStyle}>Field</th><th style={cellStyle}>Type</th></tr></thead>
              <tbody>
                {plan.fields.map(field => (
                  <tr key={field.InternalName}>
                    <td style={cellStyle}>{field.InternalName}</td>
                    <td style={cellStyle}>{field.TypeAsString}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <Checkbox
            disabled={busy}
            checked={confirmed}
            label="I reviewed the field mapping, existing target data, and workflow warning."
            onChange={(_event, checked) => setConfirmed(Boolean(checked))}
          />
        </>}

        <div className={styles.log} role="log" aria-live="polite">
          {messages.map((message, index) => <div key={`${index}-${message}`}>{message}</div>)}
        </div>
      </div>
      <Dialog
        hidden={!pendingOperation}
        onDismiss={() => setPendingOperation(undefined)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Potential production data risk',
          subText: `The destination URL does not contain "dev" or "uat": ${targetSiteUrl || '(empty URL)'}.`
        }}
        modalProps={{ isBlocking: true }}
      >
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          You are about to {pendingOperation === 'migrate' ? 'write migrated items and attachments to' : 'inspect'} a
          destination that may be a production site. Production data, workflows, notifications, and downstream
          integrations may be affected. Verify the destination URL and obtain the required approval before continuing.
        </MessageBar>
        <DialogFooter>
          <PrimaryButton
            text="I understand, continue"
            iconProps={{ iconName: 'Warning' }}
            onClick={() => confirmOperation()}
          />
          <DefaultButton
            text="Cancel"
            onClick={() => setPendingOperation(undefined)}
          />
        </DialogFooter>
      </Dialog>
      <Dialog
        hidden={!migrationSucceeded}
        onDismiss={() => setMigrationSucceeded(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Migration successful',
          subText: `${plan?.items.length || 0} records were processed.`
        }}
        modalProps={{ isBlocking: true }}
      >
        <MessageBar
          messageBarType={migrationGroupWarnings.length ? MessageBarType.warning : MessageBarType.success}
          isMultiline
        >
          {migrationGroupWarnings.length
            ? 'Migration completed, but some group/principal values were not copied. Check them separately.'
            : 'The migration completed successfully. Please verify the records in the destination list.'}
        </MessageBar>
        {migrationGroupWarnings.length > 0 && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            <ul>
              {migrationGroupWarnings.map((warning, index) => <li key={`${index}-${warning}`}>{warning}</li>)}
            </ul>
          </MessageBar>
        )}
        <DialogFooter>
          <PrimaryButton
            text="Done"
            onClick={() => setMigrationSucceeded(false)}
          />
        </DialogFooter>
      </Dialog>
      {verificationReport && <Dialog
        hidden={false}
        onDismiss={() => setVerificationReport(undefined)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Consistency check results',
          subText: `Compared ${verificationReport.sourceRecords} source records with ${
            verificationReport.targetRecords
          } destination records. ${verificationReport.matchingRecords} records match exactly.`
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: '90vw', width: 900 } }
        }}
      >
        <MessageBar
          messageBarType={
            verificationReport.groupWarnings.length > 0
              ? MessageBarType.warning
              : verificationReport.sourceOnlyRecords === 0 && verificationReport.targetOnlyRecords === 0
                ? MessageBarType.success
                : MessageBarType.warning
          }
          isMultiline
        >
          {verificationReport.groupWarnings.length > 0
            ? 'The comparable values were checked, but unresolved group/principal values were excluded.'
            : verificationReport.sourceOnlyRecords === 0 && verificationReport.targetOnlyRecords === 0
            ? 'Both lists contain the same values for all participating migration fields, including duplicate counts.'
            : 'The lists differ in the participating migration field values.'}
          {' '}Source-only records: {verificationReport.sourceOnlyRecords};
          {' '}destination-only records: {verificationReport.targetOnlyRecords}.
        </MessageBar>
        {verificationReport.groupWarnings.length > 0 && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            <strong>Some group/principal values could not be resolved and were excluded from this comparison.
              Check these values separately:</strong>
            <ul>
              {verificationReport.groupWarnings.map((warning, index) => <li key={`${index}-${warning}`}>{warning}</li>)}
            </ul>
          </MessageBar>
        )}
        {verificationReport.differences.length > 0 && (
          <div className={styles.tableWrap}>
            <table className={styles.fieldTable} aria-label="Consistency differences">
              <thead>
                <tr>
                  <th style={cellStyle}>List</th>
                  <th style={cellStyle}>Item ID</th>
                  <th style={cellStyle}>Difference</th>
                </tr>
              </thead>
              <tbody>
                {verificationReport.differences.slice(0, 50).map((difference, index) => (
                <tr key={`${difference.side}-${difference.itemId}-${index}`}>
                  <td style={cellStyle}>{difference.side}</td>
                  <td style={cellStyle}>{difference.itemId}</td>
                  <td style={cellStyle}>{difference.reason}</td>
                </tr>
                ))}
              </tbody>
            </table>
            {verificationReport.differences.length > 50 && (
              <p>Showing the first 50 of {verificationReport.differences.length} differences.</p>
            )}
          </div>
        )}
        <DialogFooter>
          <PrimaryButton text="Close" onClick={() => setVerificationReport(undefined)} />
        </DialogFooter>
      </Dialog>}
    </section>
  );
}
