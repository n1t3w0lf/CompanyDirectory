import { PeopleService } from '../../../services/PeopleService';
import { SyncService } from '../../../services/SyncService';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IPeopleDirectoryProps {
  title: string;
  description: string;
  peopleService: PeopleService;
  syncService: SyncService;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
}
