import { PeopleService } from '../../../services/PeopleService';
import { SyncService } from '../../../services/SyncService';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IPeopleDirectoryProps {
  peopleService: PeopleService;
  syncService: SyncService;
  displayMode: DisplayMode;

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

  // Font Styling
  fontSize: number;
  fontColor: string;

  // Icon Styling
  iconSize: number;
  iconColor: string;
}
