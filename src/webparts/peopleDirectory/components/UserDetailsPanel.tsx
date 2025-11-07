import * as React from 'react';
import { IUserProfile } from '../../../models/IUserProfile';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Icon } from '@fluentui/react/lib/Icon';
import { Link } from '@fluentui/react/lib/Link';
import { Separator } from '@fluentui/react/lib/Separator';
import styles from './PeopleDirectory.module.scss';

export interface IUserDetailsPanelProps {
  user: IUserProfile;
  isOpen: boolean;
  onDismiss: () => void;
}

export const UserDetailsPanel: React.FC<IUserDetailsPanelProps> = ({ user, isOpen, onDismiss }) => {
  const getInitials = (): string => {
    const firstName = user.givenName || '';
    const lastName = user.surname || '';
    return `${firstName.charAt(0)}${lastName.charAt(0)}`.toUpperCase() || user.displayName.charAt(0).toUpperCase();
  };

  const renderDetailRow = (icon: string, label: string, value?: string, isLink?: boolean, linkHref?: string): JSX.Element | null => {
    if (!value) return null;

    return (
      <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="start">
        <Icon iconName={icon} className={styles.detailIcon} />
        <Stack tokens={{ childrenGap: 2 }} grow>
          <Text variant="small" className={styles.detailLabel}>
            {label}
          </Text>
          {isLink && linkHref ? (
            <Link href={linkHref} target="_blank" className={styles.detailValue}>
              {value}
            </Link>
          ) : (
            <Text variant="medium" className={styles.detailValue}>
              {value}
            </Text>
          )}
        </Stack>
      </Stack>
    );
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Profile Details"
      closeButtonAriaLabel="Close"
      isLightDismiss
    >
      <Stack tokens={{ childrenGap: 24 }} styles={{ root: { marginTop: 20 } }}>
        {/* Profile Header */}
        <Stack horizontalAlign="center" tokens={{ childrenGap: 16 }}>
          <Persona
            imageUrl={user.photoUrl || undefined}
            imageInitials={getInitials()}
            size={PersonaSize.size120}
            hidePersonaDetails
            imageShouldFadeIn
            imageShouldStartVisible={!!user.photoUrl}
          />
          <Stack horizontalAlign="center" tokens={{ childrenGap: 4 }}>
            <Text variant="xLarge" className={styles.detailsName}>
              {user.displayName}
            </Text>
            {user.jobTitle && (
              <Text variant="medium" className={styles.detailsJobTitle}>
                {user.jobTitle}
              </Text>
            )}
          </Stack>
        </Stack>

        <Separator />

        {/* Contact Information */}
        <Stack tokens={{ childrenGap: 16 }}>
          <Text variant="large" className={styles.sectionTitle}>
            Contact Information
          </Text>

          {renderDetailRow('Mail', 'Email', user.mail, true, `mailto:${user.mail}`)}

          {user.businessPhones && user.businessPhones.length > 0 &&
            renderDetailRow('Phone', 'Business Phone', user.businessPhones[0], true, `tel:${user.businessPhones[0]}`)}

          {user.mobilePhone &&
            renderDetailRow('CellPhone', 'Mobile Phone', user.mobilePhone, true, `tel:${user.mobilePhone}`)}
        </Stack>

        <Separator />

        {/* Organization Information */}
        <Stack tokens={{ childrenGap: 16 }}>
          <Text variant="large" className={styles.sectionTitle}>
            Organization
          </Text>

          {renderDetailRow('Org', 'Department', user.department)}
          {renderDetailRow('CompanyDirectory', 'Company', user.companyName)}
          {renderDetailRow('POI', 'Office Location', user.officeLocation)}
          {renderDetailRow('CityNext', 'City', user.city)}
          {renderDetailRow('Globe', 'Country', user.country)}
          {renderDetailRow('Contact', 'Employee ID', user.employeeId)}
        </Stack>

        {user.manager && (
          <>
            <Separator />
            <Stack tokens={{ childrenGap: 16 }}>
              <Text variant="large" className={styles.sectionTitle}>
                Manager
              </Text>
              <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                <Persona
                  imageUrl={user.manager.photoUrl || undefined}
                  imageInitials={user.manager.displayName.charAt(0)}
                  size={PersonaSize.size40}
                  hidePersonaDetails
                />
                <Stack>
                  <Text variant="medium">{user.manager.displayName}</Text>
                  {user.manager.jobTitle && (
                    <Text variant="small" className={styles.userInfo}>
                      {user.manager.jobTitle}
                    </Text>
                  )}
                </Stack>
              </Stack>
            </Stack>
          </>
        )}

        {/* Metadata */}
        {user.lastVerified && (
          <>
            <Separator />
            <Stack>
              <Text variant="small" className={styles.metadata}>
                Last verified: {new Date(user.lastVerified).toLocaleString()}
              </Text>
            </Stack>
          </>
        )}
      </Stack>
    </Panel>
  );
};
