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
  onClick: () => void;

  // User Properties Visibility
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

  // Font Styling
  fontSize?: number;
  fontColor?: string;

  // Icon Styling
  iconSize?: number;
  iconColor?: string;
}

export const UserCard: React.FC<IUserCardProps> = ({
  user,
  onClick,
  showEmail = true,
  showJobTitle = true,
  showDepartment = true,
  showOfficeLocation = true,
  showBusinessPhones = true,
  showMobilePhone = true,
  showCity = true,
  showCountry = true,
  showCompanyName = true,
  showEmployeeId = true,
  fontSize = 14,
  fontColor = '#323130',
  iconSize = 24,
  iconColor = '#0078d4'
}) => {
  const getInitials = (): string => {
    const firstName = user.givenName || '';
    const lastName = user.surname || '';
    return `${firstName.charAt(0)}${lastName.charAt(0)}`.toUpperCase() || user.displayName.charAt(0).toUpperCase();
  };

  const textStyle: React.CSSProperties = {
    color: fontColor,
    fontSize: `${fontSize}px`
  };

  const iconStyle: React.CSSProperties = {
    fontSize: `${iconSize}px`,
    color: iconColor
  };

  return (
    <div className={styles.userCard} onClick={onClick} role="button" tabIndex={0}>
      <Stack tokens={{ childrenGap: 12 }} horizontal>
        <Stack.Item>
          <Persona
            imageUrl={user.photoUrl || undefined}
            imageInitials={getInitials()}
            size={PersonaSize.size72}
            hidePersonaDetails
            imageShouldFadeIn
            imageShouldStartVisible={!!user.photoUrl}
          />
        </Stack.Item>

        <Stack tokens={{ childrenGap: 8 }} grow>
          <Text variant="large" block className={styles.userName} style={textStyle}>
            {user.displayName}
          </Text>

          {showJobTitle && user.jobTitle && (
            <Text variant="medium" block className={styles.userJobTitle} style={textStyle}>
              {user.jobTitle}
            </Text>
          )}

          {showDepartment && user.department && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon iconName="Org" className={styles.icon} style={iconStyle} />
              <Text variant="small" className={styles.userInfo} style={textStyle}>
                {user.department}
              </Text>
            </Stack>
          )}

          {showOfficeLocation && user.officeLocation && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon iconName="POI" className={styles.icon} style={iconStyle} />
              <Text variant="small" className={styles.userInfo} style={textStyle}>
                {user.officeLocation}
              </Text>
            </Stack>
          )}

          {showCity && user.city && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon iconName="CityNext" className={styles.icon} style={iconStyle} />
              <Text variant="small" className={styles.userInfo} style={textStyle}>
                {user.city}
              </Text>
            </Stack>
          )}

          {showCountry && user.country && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon iconName="Globe" className={styles.icon} style={iconStyle} />
              <Text variant="small" className={styles.userInfo} style={textStyle}>
                {user.country}
              </Text>
            </Stack>
          )}

          {showCompanyName && user.companyName && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon iconName="CompanyDirectory" className={styles.icon} style={iconStyle} />
              <Text variant="small" className={styles.userInfo} style={textStyle}>
                {user.companyName}
              </Text>
            </Stack>
          )}

          {showEmployeeId && user.employeeId && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon iconName="Contact" className={styles.icon} style={iconStyle} />
              <Text variant="small" className={styles.userInfo} style={textStyle}>
                ID: {user.employeeId}
              </Text>
            </Stack>
          )}
          {/* Action Buttons */}
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 8 }}>
            {showEmail && user.mail && (
              <IconButton
                iconProps={{ iconName: 'Mail' }}
                title={`Email ${user.displayName}`}
                ariaLabel="Send email"
                onClick={(e) => {
                  e.stopPropagation();
                  window.location.href = `mailto:${user.mail}`;
                }}
                className={styles.actionButton}
                style={{ color: iconColor }}
              />
            )}
            {showBusinessPhones && (user.businessPhones && user.businessPhones.length > 0) && (
              <IconButton
                iconProps={{ iconName: 'Phone' }}
                title={`Call ${user.displayName}`}
                ariaLabel="Call"
                onClick={(e) => {
                  e.stopPropagation();
                  window.location.href = `tel:${user.businessPhones[0]}`;
                }}
                className={styles.actionButton}
                style={{ color: iconColor }}
              />
            )}
            {showMobilePhone && user.mobilePhone && (
              <IconButton
                iconProps={{ iconName: 'CellPhone' }}
                title={`Call mobile ${user.displayName}`}
                ariaLabel="Call mobile"
                onClick={(e) => {
                  e.stopPropagation();
                  window.location.href = `tel:${user.mobilePhone}`;
                }}
                className={styles.actionButton}
                style={{ color: iconColor }}
              />
            )}
          </Stack>
        </Stack>
      </Stack>
    </div>
  );
};
