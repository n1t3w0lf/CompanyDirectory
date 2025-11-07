import * as React from 'react';
import { IUserProfile } from '../../../models/IUserProfile';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Icon } from '@fluentui/react/lib/Icon';
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

  // Profile Name Styling
  profileNameFontSize?: number;
  profileNameFontColor?: string;

  // Job Title Styling
  jobTitleFontSize?: number;
  jobTitleFontColor?: string;
  jobTitleBold?: boolean;

  // Profile Properties Styling
  profilePropertiesFontSize?: number;
  profilePropertiesFontColor?: string;

  // Property Display Order
  propertyDisplayOrder?: string;

  // Text Truncation
  textEllipsisLength?: number;

  // Icon Styling
  iconSize?: number;
  iconColor?: string;

  // Profile Card Background
  profileCardBackgroundColor?: string;
  profileCardBackgroundImage?: string;

  // Profile Picture Visibility
  showProfilePicture?: boolean;
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
  profileNameFontSize = 18,
  profileNameFontColor = '#323130',
  jobTitleFontSize = 16,
  jobTitleFontColor = '#323130',
  jobTitleBold = true,
  profilePropertiesFontSize = 14,
  profilePropertiesFontColor = '#605E5C',
  propertyDisplayOrder = 'jobTitle,email,department,officeLocation,city,country,companyName,businessPhones,mobilePhone,employeeId',
  textEllipsisLength = 50,
  iconSize = 20,
  iconColor = '#0078d4',
  profileCardBackgroundColor = '',
  profileCardBackgroundImage = '',
  showProfilePicture = true
}) => {
  const getInitials = (): string => {
    const firstName = user.givenName || '';
    const lastName = user.surname || '';
    return `${firstName.charAt(0)}${lastName.charAt(0)}`.toUpperCase() || user.displayName.charAt(0).toUpperCase();
  };

  const truncateText = (text: string, maxLength: number): string => {
    if (!text || text.length <= maxLength) {
      return text;
    }
    return text.substring(0, maxLength) + '...';
  };

  const nameStyle: React.CSSProperties = {
    color: profileNameFontColor,
    fontSize: `${profileNameFontSize}px`,
    fontWeight: 600
  };

  const jobTitleStyle: React.CSSProperties = {
    color: jobTitleFontColor,
    fontSize: `${jobTitleFontSize}px`,
    fontWeight: jobTitleBold ? 600 : 400
  };

  const propertiesStyle: React.CSSProperties = {
    color: profilePropertiesFontColor,
    fontSize: `${profilePropertiesFontSize}px`
  };

  const iconStyle: React.CSSProperties = {
    fontSize: `${iconSize}px`,
    color: iconColor
  };

  // Build card background style
  const cardBackgroundStyle: React.CSSProperties = {};
  if (profileCardBackgroundImage) {
    cardBackgroundStyle.backgroundImage = `url(${profileCardBackgroundImage})`;
    cardBackgroundStyle.backgroundSize = 'cover';
    cardBackgroundStyle.backgroundPosition = 'center';
    cardBackgroundStyle.backgroundRepeat = 'no-repeat';
  } else if (profileCardBackgroundColor) {
    cardBackgroundStyle.backgroundColor = profileCardBackgroundColor;
  }

  // Map of field renderers
  const fieldRenderers: { [key: string]: () => JSX.Element | null } = {
    jobTitle: () => showJobTitle ? (
      <Text variant="medium" block className={styles.userJobTitle} style={jobTitleStyle} title={user.jobTitle || ''}>
        {truncateText(user.jobTitle || '', textEllipsisLength)}
      </Text>
    ) : null,
    email: () => showEmail ? (
      <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
        <Icon iconName="Mail" className={styles.icon} style={iconStyle} />
        <Text variant="small" className={styles.userInfo} style={propertiesStyle} title={user.mail || ''}>
          {truncateText(user.mail || '', textEllipsisLength)}
        </Text>
      </Stack>
    ) : null,
    department: () => showDepartment ? (
      <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
        <Icon iconName="Org" className={styles.icon} style={iconStyle} />
        <Text variant="small" className={styles.userInfo} style={propertiesStyle} title={user.department || ''}>
          {truncateText(user.department || '', textEllipsisLength)}
        </Text>
      </Stack>
    ) : null,
    officeLocation: () => showOfficeLocation ? (
      <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
        <Icon iconName="POI" className={styles.icon} style={iconStyle} />
        <Text variant="small" className={styles.userInfo} style={propertiesStyle} title={user.officeLocation || ''}>
          {truncateText(user.officeLocation || '', textEllipsisLength)}
        </Text>
      </Stack>
    ) : null,
    city: () => showCity ? (
      <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
        <Icon iconName="CityNext" className={styles.icon} style={iconStyle} />
        <Text variant="small" className={styles.userInfo} style={propertiesStyle} title={user.city || ''}>
          {truncateText(user.city || '', textEllipsisLength)}
        </Text>
      </Stack>
    ) : null,
    country: () => showCountry ? (
      <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
        <Icon iconName="Globe" className={styles.icon} style={iconStyle} />
        <Text variant="small" className={styles.userInfo} style={propertiesStyle} title={user.country || ''}>
          {truncateText(user.country || '', textEllipsisLength)}
        </Text>
      </Stack>
    ) : null,
    companyName: () => showCompanyName ? (
      <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
        <Icon iconName="CompanyDirectory" className={styles.icon} style={iconStyle} />
        <Text variant="small" className={styles.userInfo} style={propertiesStyle} title={user.companyName || ''}>
          {truncateText(user.companyName || '', textEllipsisLength)}
        </Text>
      </Stack>
    ) : null,
    businessPhones: () => showBusinessPhones ? (
      <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
        <Icon iconName="Phone" className={styles.icon} style={iconStyle} />
        <Text variant="small" className={styles.userInfo} style={propertiesStyle} title={(user.businessPhones && user.businessPhones[0]) || ''}>
          {truncateText((user.businessPhones && user.businessPhones[0]) || '', textEllipsisLength)}
        </Text>
      </Stack>
    ) : null,
    mobilePhone: () => showMobilePhone ? (
      <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
        <Icon iconName="CellPhone" className={styles.icon} style={iconStyle} />
        <Text variant="small" className={styles.userInfo} style={propertiesStyle} title={user.mobilePhone || ''}>
          {truncateText(user.mobilePhone || '', textEllipsisLength)}
        </Text>
      </Stack>
    ) : null,
    employeeId: () => showEmployeeId ? (
      <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
        <Icon iconName="Contact" className={styles.icon} style={iconStyle} />
        <Text variant="small" className={styles.userInfo} style={propertiesStyle} title={user.employeeId ? `ID: ${user.employeeId}` : ''}>
          ID: {truncateText(user.employeeId || '', textEllipsisLength)}
        </Text>
      </Stack>
    ) : null
  };

  // Parse display order and render fields
  const fieldOrder = propertyDisplayOrder.split(',').map(f => f.trim());

  return (
    <div className={styles.userCard} onClick={onClick} role="button" tabIndex={0} style={cardBackgroundStyle}>
      <Stack tokens={{ childrenGap: 12 }} horizontal>
        {showProfilePicture && (
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
        )}

        <Stack tokens={{ childrenGap: 8 }} grow>
          <Text variant="large" block className={styles.userName} style={nameStyle} title={user.displayName}>
            {truncateText(user.displayName, textEllipsisLength)}
          </Text>

          {fieldOrder.map((fieldName, index) => {
            const renderer = fieldRenderers[fieldName];
            return renderer ? <React.Fragment key={`${fieldName}-${index}`}>{renderer()}</React.Fragment> : null;
          })}
        </Stack>
      </Stack>
    </div>
  );
};
