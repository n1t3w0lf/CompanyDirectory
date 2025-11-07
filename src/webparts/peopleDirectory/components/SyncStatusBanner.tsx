import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { ISyncStatus } from '../../../services/SyncService';
import styles from './PeopleDirectory.module.scss';

export interface ISyncStatusBannerProps {
  syncStatus: ISyncStatus | null;
  onStartSync: () => void;
  onCancelSync: () => void;
}

export const SyncStatusBanner: React.FC<ISyncStatusBannerProps> = ({ syncStatus, onStartSync, onCancelSync }) => {
  if (!syncStatus) {
    return null;
  }

  // Show sync needed message
  if (!syncStatus.isRunning && !syncStatus.lastSyncDate) {
    return (
      <MessageBar
        messageBarType={MessageBarType.warning}
        actions={
          <PrimaryButton onClick={onStartSync}>
            Start Initial Sync
          </PrimaryButton>
        }
      >
        <strong>Initial sync required:</strong> Click &quot;Start Initial Sync&quot; to populate the directory with all users from Entra ID.
        This is a one-time operation that will fetch all {syncStatus.totalUsers > 0 ? syncStatus.totalUsers : ''} users.
      </MessageBar>
    );
  }

  // Show sync in progress
  if (syncStatus.isRunning) {
    return (
      <Stack tokens={{ childrenGap: 12 }} className={styles.syncStatusContainer}>
        <MessageBar
          messageBarType={MessageBarType.info}
          actions={
            <DefaultButton onClick={onCancelSync}>
              Cancel
            </DefaultButton>
          }
        >
          <strong>Synchronization in progress...</strong>
          {syncStatus.currentBatch > 0 && (
            <Text>
              {' '}Processing batch {syncStatus.currentBatch} of {syncStatus.totalBatches}
              ({syncStatus.processedUsers} / {syncStatus.totalUsers} users)
            </Text>
          )}
        </MessageBar>
        <ProgressIndicator
          label="Syncing users from Entra ID"
          description={`${syncStatus.progress}% complete`}
          percentComplete={syncStatus.progress / 100}
        />
      </Stack>
    );
  }

  // Show sync complete
  if (syncStatus.lastSyncDate) {
    const hasErrors = syncStatus.errors && syncStatus.errors.length > 0;
    return (
      <MessageBar
        messageBarType={hasErrors ? MessageBarType.warning : MessageBarType.success}
        isMultiline={hasErrors}
      >
        <strong>Last sync:</strong> {syncStatus.lastSyncDate.toLocaleString()}
        {' '}({syncStatus.totalUsers} users)
        {hasErrors && (
          <>
            <br />
            <Text variant="small">
              {syncStatus.errors.length} errors occurred during sync. Check console for details.
            </Text>
          </>
        )}
      </MessageBar>
    );
  }

  return null;
};
