import { PeopleService } from '../../../services/PeopleService';
import { SyncService } from '../../../services/SyncService';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IPeopleDirectoryProps {
  peopleService: PeopleService;
  syncService: SyncService;
  displayMode: DisplayMode;

  // Header Configuration
  headingText: string;
  headingFontSize: number;
  headingFontColor: string;
  subtextText: string;
  subtextFontSize: number;
  subtextFontColor: string;

  // User Properties to Display
  showEmail: boolean;
  showJobTitle: boolean;
  showDepartment: boolean;
  showOfficeLocation: boolean;
  showBusinessPhones: boolean;
  showMobilePhone: boolean;
  showCity: boolean;
  showCountry: boolean;
  showCompanyName: boolean;
  showEmployeeId: boolean;

  // Profile Name Styling
  profileNameFontSize: number;
  profileNameFontColor: string;

  // Profile Properties Styling
  profilePropertiesFontSize: number;
  profilePropertiesFontColor: string;

  // Icon Styling
  iconSize: number;
  iconColor: string;
}
