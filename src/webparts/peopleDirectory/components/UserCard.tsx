import * as React from 'react';
import { IUserProfile } from '../../../models/IUserProfile';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import styles from './PeopleDirectory.module.scss';

export interface IUserCardProps {
  user: IUserProfile;
  onUserClick: (user: IUserProfile) => void;
}

export const UserCard: React.FC<IUserCardProps> = ({ user, onUserClick }) => {
  const handleClick = React.useCallback((): void => {
    onUserClick(user);
  }, [user, onUserClick]);

  const handleEmailClick = React.useCallback((e: React.MouseEvent<HTMLButtonElement>): void => {
    e.stopPropagation();
    if (user.mail) {
      window.location.href = `mailto:${user.mail}`;
    }
  }, [user.mail]);

  const handleBusinessPhoneClick = React.useCallback((e: React.MouseEvent<HTMLButtonElement>): void => {
    e.stopPropagation();
    if (user.businessPhones && user.businessPhones.length > 0) {
      window.location.href = `tel:${user.businessPhones[0]}`;
    }
  }, [user.businessPhones]);

  const handleMobilePhoneClick = React.useCallback((e: React.MouseEvent<HTMLButtonElement>): void => {
    e.stopPropagation();
    if (user.mobilePhone) {
      window.location.href = `tel:${user.mobilePhone}`;
    }
  }, [user.mobilePhone]);

  const getInitials = (): string => {
    const firstName = user.givenName || '';
    const lastName = user.surname || '';
    return `${firstName.charAt(0)}${lastName.charAt(0)}`.toUpperCase() || user.displayName.charAt(0).toUpperCase();
  };

  return (
    <div className={styles.userCard} onClick={handleClick} role="button" tabIndex={0}>
      <Stack tokens={{ childrenGap: 12 }}>
        <Stack horizontal horizontalAlign="center">
          <Persona
            imageUrl={user.photoUrl || undefined}
            imageInitials={getInitials()}
            size={PersonaSize.size72}
            hidePersonaDetails
            imageShouldFadeIn
            imageShouldStartVisible={!!user.photoUrl}
          />
        </Stack>

        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="large" block className={styles.userName}>
            {user.displayName}
          </Text>

          {user.jobTitle && (
            <Text variant="medium" block className={styles.userJobTitle}>
              {user.jobTitle}
            </Text>
          )}

          {user.department && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon iconName="Org" className={styles.icon} />
              <Text variant="small" className={styles.userInfo}>
                {user.department}
              </Text>
            </Stack>
          )}

          {user.officeLocation && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon iconName="POI" className={styles.icon} />
              <Text variant="small" className={styles.userInfo}>
                {user.officeLocation}
              </Text>
            </Stack>
          )}
        </Stack>

        <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 8 }}>
          {user.mail && (
            <IconButton
              iconProps={{ iconName: 'Mail' }}
              title={`Email ${user.displayName}`}
              ariaLabel="Send email"
              onClick={handleEmailClick}
              className={styles.actionButton}
            />
          )}
          {(user.businessPhones && user.businessPhones.length > 0) && (
            <IconButton
              iconProps={{ iconName: 'Phone' }}
              title={`Call ${user.displayName}`}
              ariaLabel="Call"
              onClick={handleBusinessPhoneClick}
              className={styles.actionButton}
            />
          )}
          {user.mobilePhone && (
            <IconButton
              iconProps={{ iconName: 'CellPhone' }}
              title={`Call mobile ${user.displayName}`}
              ariaLabel="Call mobile"
              onClick={handleMobilePhoneClick}
              className={styles.actionButton}
            />
          )}
        </Stack>
      </Stack>
    </div>
  );
};
