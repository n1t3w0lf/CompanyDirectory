import { PeopleService } from '../../../services/PeopleService';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IPeopleDirectoryProps {
  title: string;
  description: string;
  peopleService: PeopleService;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
}
