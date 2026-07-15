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

  // Styling props
  profileNameFontSize?: number;
  profileNameFontColor?: string;
  jobTitleFontSize?: number;
  jobTitleFontColor?: string;
  jobTitleBold?: boolean;
  profilePropertiesFontSize?: number;
  profilePropertiesFontColor?: string;
  iconSize?: number;
  iconColor?: string;
  showProfilePicture?: boolean;

  // Visibility props
  showEmail?: boolean;
  showJobTitle?: boolean;
  showDepartment?: boolean;
  showOfficeLocation?: boolean;
  showBusinessPhones?: boolean;
  showMobilePhone?: boolean;
  showCity?: boolean;
  showCountry?: boolean;
  showCompanyName?: boolean;
  showEmployeeId?: boolean;
}

export const UserDetailsPanel: React.FC<IUserDetailsPanelProps> = ({
  user,
  isOpen,
  onDismiss,
  profileNameFontSize = 24,
  profileNameFontColor = '#323130',
  jobTitleFontSize = 16,
  jobTitleFontColor = '#323130',
  jobTitleBold = true,
  profilePropertiesFontSize = 14,
  profilePropertiesFontColor = '#605E5C',
  iconSize = 20,
  iconColor = '#0078d4',
  showProfilePicture = true,
  showEmail = true,
  showJobTitle = true,
  showDepartment = true,
  showOfficeLocation = true,
  showBusinessPhones = true,
  showMobilePhone = true,
  showCity = true,
  showCountry = true,
  showCompanyName = true,
  showEmployeeId = true
}) => {
  const getInitials = (): string => {
    const firstName = user.givenName || '';
    const lastName = user.surname || '';
    return `${firstName.charAt(0)}${lastName.charAt(0)}`.toUpperCase() || user.displayName.charAt(0).toUpperCase();
  };

  const renderDetailRow = (icon: string, label: string, value?: string, isLink?: boolean, linkHref?: string): JSX.Element | null => {
    if (!value) return null;

    return (
      <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="start">
        <Icon
          iconName={icon}
          className={styles.detailIcon}
          style={{ fontSize: `${iconSize}px`, color: iconColor }}
        />
        <Stack tokens={{ childrenGap: 2 }} grow>
          <Text variant="small" className={styles.detailLabel}>
            {label}
          </Text>
          {isLink && linkHref ? (
            <Link
              href={linkHref}
              target="_blank"
              className={styles.detailValue}
              style={{
                fontSize: `${profilePropertiesFontSize}px`,
                color: profilePropertiesFontColor
              }}
            >
              {value}
            </Link>
          ) : (
            <Text
              variant="medium"
              className={styles.detailValue}
              style={{
                fontSize: `${profilePropertiesFontSize}px`,
                color: profilePropertiesFontColor
              }}
            >
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
      styles={{ closeButton: { selectors: { '@media (max-width: 640px)': { width: 44, height: 44 } } } }}
    >
      <Stack tokens={{ childrenGap: 24 }} styles={{ root: { marginTop: 20 } }}>
        {/* Profile Header */}
        <Stack horizontalAlign="center" tokens={{ childrenGap: 16 }}>
          {showProfilePicture && (
            <Persona
              imageUrl={user.photoUrl || undefined}
              imageInitials={getInitials()}
              size={PersonaSize.size120}
              hidePersonaDetails
              imageShouldFadeIn
              imageShouldStartVisible={!!user.photoUrl}
            />
          )}
          <Stack horizontalAlign="center" tokens={{ childrenGap: 4 }}>
            <Text
              variant="xLarge"
              className={styles.detailsName}
              style={{
                fontSize: `${profileNameFontSize}px`,
                color: profileNameFontColor
              }}
            >
              {user.displayName}
            </Text>
            {showJobTitle && user.jobTitle && (
              <Text
                variant="medium"
                className={styles.detailsJobTitle}
                style={{
                  fontSize: `${jobTitleFontSize}px`,
                  color: jobTitleFontColor,
                  fontWeight: jobTitleBold ? 'bold' : 'normal'
                }}
              >
                {user.jobTitle}
              </Text>
            )}
          </Stack>
        </Stack>

        {(showEmail || showBusinessPhones || showMobilePhone) && <Separator />}

        {/* Contact Information */}
        {(showEmail || showBusinessPhones || showMobilePhone) && (
          <Stack tokens={{ childrenGap: 16 }}>
            <Text variant="large" className={styles.sectionTitle}>
              Contact Information
            </Text>

            {showEmail && renderDetailRow('Mail', 'Email', user.mail, true, `mailto:${user.mail}`)}

            {showBusinessPhones && user.businessPhones && user.businessPhones.length > 0 &&
              renderDetailRow('Phone', 'Business Phone', user.businessPhones[0], true, `tel:${user.businessPhones[0]}`)}

            {showMobilePhone && user.mobilePhone &&
              renderDetailRow('CellPhone', 'Mobile Phone', user.mobilePhone, true, `tel:${user.mobilePhone}`)}
          </Stack>
        )}

        {(showDepartment || showCompanyName || showOfficeLocation || showCity || showCountry || showEmployeeId) && <Separator />}

        {/* Organization Information */}
        {(showDepartment || showCompanyName || showOfficeLocation || showCity || showCountry || showEmployeeId) && (
          <Stack tokens={{ childrenGap: 16 }}>
            <Text variant="large" className={styles.sectionTitle}>
              Organization
            </Text>

            {showDepartment && renderDetailRow('Org', 'Department', user.department)}
            {showCompanyName && renderDetailRow('CompanyDirectory', 'Company', user.companyName)}
            {showOfficeLocation && renderDetailRow('POI', 'Office Location', user.officeLocation)}
            {showCity && renderDetailRow('CityNext', 'City', user.city)}
            {showCountry && renderDetailRow('Globe', 'Country', user.country)}
            {showEmployeeId && renderDetailRow('Contact', 'Employee ID', user.employeeId)}
          </Stack>
        )}

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
