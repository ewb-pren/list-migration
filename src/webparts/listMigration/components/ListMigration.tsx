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
  IListMigrationPlan,
  IListMigrationRequest
} from '../services/ListMigrationService';

export default function ListMigration(props: IListMigrationProps): React.ReactElement {
  const currentSiteUrl = props.context.pageContext.web.absoluteUrl;
  const [sourceSiteUrl, setSourceSiteUrl] = React.useState(currentSiteUrl);
  const [targetSiteUrl, setTargetSiteUrl] = React.useState(currentSiteUrl);
  const [sourceListTitle, setSourceListTitle] = React.useState('');
  const [targetListTitle, setTargetListTitle] = React.useState('');
  const [plan, setPlan] = React.useState<IListMigrationPlan>();
  const [busy, setBusy] = React.useState(false);
  const [confirmed, setConfirmed] = React.useState(false);
  const [error, setError] = React.useState('');
  const [messages, setMessages] = React.useState<string[]>([]);
  const [pendingOperation, setPendingOperation] = React.useState<'preflight' | 'migrate'>();
  const inFlight = React.useRef(false);
  const releaseInFlight = (): void => {
    inFlight.current = false;
  };

  const resetPlan = (): void => {
    setPlan(undefined);
    setConfirmed(false);
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
      setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      releaseInFlight();
      setBusy(false);
    }
  };

  const runMigration = async (): Promise<void> => {
    if (!plan || !confirmed || inFlight.current) return;
    inFlight.current = true;
    setBusy(true);
    setError('');
    try {
      const service = new ListMigrationService(props.context, getRequest());
      await service.migrate(plan, message => setMessages(previous => [...previous.slice(-29), message]));
      setConfirmed(false);
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : String(failure));
    } finally {
      releaseInFlight();
      setBusy(false);
    }
  };

  const requestOperation = async (operation: 'preflight' | 'migrate'): Promise<void> => {
    if (isProductionRisk) {
      setPendingOperation(operation);
      return;
    }
    if (operation === 'preflight') {
      await runPreflight();
    } else {
      await runMigration();
    }
  };

  const confirmOperation = async (): Promise<void> => {
    const operation = pendingOperation;
    setPendingOperation(undefined);
    if (operation === 'preflight') {
      await runPreflight();
    } else if (operation === 'migrate') {
      await runMigration();
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
          {busy && <Spinner label="Processing migration" />}
        </Stack>

        {error && <MessageBar messageBarType={MessageBarType.error} role="alert">{error}</MessageBar>}

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
    </section>
  );
}
