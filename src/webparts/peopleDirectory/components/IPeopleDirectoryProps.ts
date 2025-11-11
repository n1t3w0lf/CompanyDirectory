import { PeopleService } from "../../../services/PeopleService";
import { SyncService } from "../../../services/SyncService";
import { DisplayMode } from "@microsoft/sp-core-library";

export interface IPeopleDirectoryProps {
  // Services
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

  // Job Title Styling
  jobTitleFontSize: number;
  jobTitleFontColor: string;
  jobTitleBold: boolean;

  // Profile Properties Styling
  profilePropertiesFontSize: number;
  profilePropertiesFontColor: string;

  // Property Display Order
  propertyDisplayOrder: string;

  // Text Truncation
  textEllipsisLength: number;

  // Icon Styling
  iconSize: number;
  iconColor: string;

  // Profile Card Background
  profileCardBackgroundColor: string;
  profileCardBackgroundImage: string;

  // Webpart Background
  webpartBackgroundColor: string;
  webpartBackgroundImage: string;

  // Letter Index
  showLetterIndex: boolean;

  // Profile Picture
  showProfilePicture: boolean;

  // Search Button Configuration
  searchButtonText: string;
  searchButtonColor: string;
  searchButtonTextSize: number;
  searchButtonHoverColor: string;
}
