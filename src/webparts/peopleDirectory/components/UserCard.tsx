import * as React from 'react';
import { IUserProfile } from '../../../models/IUserProfile';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './PeopleDirectory.module.scss';

export interface IUserCardProps {
  user: IUserProfile;
  onUserClick: (user: IUserProfile) => void;

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
  profileCardBackgroundColor?: string;
  profileCardBackgroundImage?: string;
  showProfilePicture?: boolean;
  textEllipsisLength?: number;

  // Visibility props
  showJobTitle?: boolean;
  showDepartment?: boolean;
  showOfficeLocation?: boolean;
  showEmail?: boolean;
  showBusinessPhones?: boolean;
  showMobilePhone?: boolean;
  showCity?: boolean;
  showCountry?: boolean;
  showCompanyName?: boolean;
  showEmployeeId?: boolean;
}

export const UserCard: React.FC<IUserCardProps> = ({
  user,
  onUserClick,
  profileNameFontSize = 18,
  profileNameFontColor = '#323130',
  jobTitleFontSize = 16,
  jobTitleFontColor = '#323130',
  jobTitleBold = true,
  profilePropertiesFontSize = 14,
  profilePropertiesFontColor = '#605E5C',
  iconSize = 20,
  iconColor = '#0078d4',
  profileCardBackgroundColor,
  profileCardBackgroundImage,
  showProfilePicture = true,
  textEllipsisLength = 50,
  showJobTitle = true,
  showDepartment = true,
  showOfficeLocation = true,
  showEmail = true,
  showBusinessPhones = true,
  showMobilePhone = true,
  showCity = true,
  showCountry = true,
  showCompanyName = true,
  showEmployeeId = true
}) => {
  const handleClick = React.useCallback((): void => {
    onUserClick(user);
  }, [user, onUserClick]);

  const getInitials = (): string => {
    const firstName = user.givenName || '';
    const lastName = user.surname || '';
    return `${firstName.charAt(0)}${lastName.charAt(0)}`.toUpperCase() || user.displayName.charAt(0).toUpperCase();
  };

  const truncateText = (text: string | undefined): string => {
    if (!text) return '';
    if (textEllipsisLength <= 0 || text.length <= textEllipsisLength) {
      return text;
    }
    return text.substring(0, textEllipsisLength) + '...';
  };

  // Build card style
  const cardStyle: React.CSSProperties = React.useMemo(() => {
    const style: React.CSSProperties = {
      ...(profileCardBackgroundColor && { backgroundColor: profileCardBackgroundColor }),
      ...(profileCardBackgroundImage && {
        backgroundImage: `url("${profileCardBackgroundImage}")`,
        backgroundSize: 'cover',
        backgroundPosition: 'center',
        backgroundRepeat: 'no-repeat'
      })
    };

    // Debug logging (only first card)
    if ((profileCardBackgroundImage || profileCardBackgroundColor) && user.displayName) {
      console.log('Card background config for', user.displayName, {
        backgroundColor: profileCardBackgroundColor,
        backgroundImage: profileCardBackgroundImage,
        styleApplied: style
      });
    }

    return style;
  }, [profileCardBackgroundColor, profileCardBackgroundImage, user.displayName]);

  return (
    <div className={styles.userCard} onClick={handleClick} role="button" tabIndex={0} style={cardStyle}>
      <Stack tokens={{ childrenGap: 12 }}>
        {/* Profile Header - Picture and Name/Title horizontal */}
        <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
          {showProfilePicture && (
            <Persona
              imageUrl={user.photoUrl || undefined}
              imageInitials={getInitials()}
              size={PersonaSize.size72}
              hidePersonaDetails
              imageShouldFadeIn
              imageShouldStartVisible={!!user.photoUrl}
            />
          )}
          <Stack tokens={{ childrenGap: 4 }} grow styles={{ root: { minWidth: 0 } }}>
            <Text
              variant="large"
              block
              className={styles.userName}
              style={{
                fontSize: `${profileNameFontSize}px`,
                color: profileNameFontColor
              }}
              title={user.displayName}
            >
              {truncateText(user.displayName)}
            </Text>

            {showJobTitle && user.jobTitle && (
              <Text
                variant="medium"
                block
                className={styles.userJobTitle}
                style={{
                  fontSize: `${jobTitleFontSize}px`,
                  color: jobTitleFontColor,
                  fontWeight: jobTitleBold ? 'bold' : 'normal'
                }}
                title={user.jobTitle}
              >
                {truncateText(user.jobTitle)}
              </Text>
            )}
          </Stack>
        </Stack>

        {/* User Properties */}
        <Stack tokens={{ childrenGap: 4 }}>

          {showEmail && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon
                iconName="Mail"
                className={styles.icon}
                style={{ fontSize: `${iconSize}px`, color: iconColor }}
              />
              <Text
                variant="small"
                className={styles.userInfo}
                style={{
                  fontSize: `${profilePropertiesFontSize}px`,
                  color: profilePropertiesFontColor
                }}
                title={user.mail || ''}
              >
                {truncateText(user.mail || '')}
              </Text>
            </Stack>
          )}

          {showBusinessPhones && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon
                iconName="Phone"
                className={styles.icon}
                style={{ fontSize: `${iconSize}px`, color: iconColor }}
              />
              <Text
                variant="small"
                className={styles.userInfo}
                style={{
                  fontSize: `${profilePropertiesFontSize}px`,
                  color: profilePropertiesFontColor
                }}
                title={user.businessPhones?.[0] || ''}
              >
                {truncateText(user.businessPhones?.[0] || '')}
              </Text>
            </Stack>
          )}

          {showMobilePhone && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon
                iconName="CellPhone"
                className={styles.icon}
                style={{ fontSize: `${iconSize}px`, color: iconColor }}
              />
              <Text
                variant="small"
                className={styles.userInfo}
                style={{
                  fontSize: `${profilePropertiesFontSize}px`,
                  color: profilePropertiesFontColor
                }}
                title={user.mobilePhone || ''}
              >
                {truncateText(user.mobilePhone || '')}
              </Text>
            </Stack>
          )}

          {showDepartment && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon
                iconName="Org"
                className={styles.icon}
                style={{ fontSize: `${iconSize}px`, color: iconColor }}
              />
              <Text
                variant="small"
                className={styles.userInfo}
                style={{
                  fontSize: `${profilePropertiesFontSize}px`,
                  color: profilePropertiesFontColor
                }}
                title={user.department || ''}
              >
                {truncateText(user.department || '')}
              </Text>
            </Stack>
          )}

          {showOfficeLocation && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon
                iconName="POI"
                className={styles.icon}
                style={{ fontSize: `${iconSize}px`, color: iconColor }}
              />
              <Text
                variant="small"
                className={styles.userInfo}
                style={{
                  fontSize: `${profilePropertiesFontSize}px`,
                  color: profilePropertiesFontColor
                }}
                title={user.officeLocation || ''}
              >
                {truncateText(user.officeLocation || '')}
              </Text>
            </Stack>
          )}

          {showCity && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon
                iconName="CityNext"
                className={styles.icon}
                style={{ fontSize: `${iconSize}px`, color: iconColor }}
              />
              <Text
                variant="small"
                className={styles.userInfo}
                style={{
                  fontSize: `${profilePropertiesFontSize}px`,
                  color: profilePropertiesFontColor
                }}
                title={user.city || ''}
              >
                {truncateText(user.city || '')}
              </Text>
            </Stack>
          )}

          {showCountry && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon
                iconName="Globe"
                className={styles.icon}
                style={{ fontSize: `${iconSize}px`, color: iconColor }}
              />
              <Text
                variant="small"
                className={styles.userInfo}
                style={{
                  fontSize: `${profilePropertiesFontSize}px`,
                  color: profilePropertiesFontColor
                }}
                title={user.country || ''}
              >
                {truncateText(user.country || '')}
              </Text>
            </Stack>
          )}

          {showCompanyName && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon
                iconName="CompanyDirectory"
                className={styles.icon}
                style={{ fontSize: `${iconSize}px`, color: iconColor }}
              />
              <Text
                variant="small"
                className={styles.userInfo}
                style={{
                  fontSize: `${profilePropertiesFontSize}px`,
                  color: profilePropertiesFontColor
                }}
                title={user.companyName || ''}
              >
                {truncateText(user.companyName || '')}
              </Text>
            </Stack>
          )}

          {showEmployeeId && (
            <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
              <Icon
                iconName="Contact"
                className={styles.icon}
                style={{ fontSize: `${iconSize}px`, color: iconColor }}
              />
              <Text
                variant="small"
                className={styles.userInfo}
                style={{
                  fontSize: `${profilePropertiesFontSize}px`,
                  color: profilePropertiesFontColor
                }}
                title={user.employeeId || ''}
              >
                {truncateText(user.employeeId || '')}
              </Text>
            </Stack>
          )}
        </Stack>
      </Stack>
    </div>
  );
};
