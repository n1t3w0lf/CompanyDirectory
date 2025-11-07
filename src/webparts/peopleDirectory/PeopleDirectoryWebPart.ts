import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'PeopleDirectoryWebPartStrings';
import { PeopleDirectory } from './components/PeopleDirectory';
import { IPeopleDirectoryProps } from './components/IPeopleDirectoryProps';
import { GraphService } from '../../services/GraphService';
import { ListService } from '../../services/ListService';
import { PeopleService } from '../../services/PeopleService';
import { SyncService } from '../../services/SyncService';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export interface IPeopleDirectoryWebPartProps {
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

export default class PeopleDirectoryWebPart extends BaseClientSideWebPart<IPeopleDirectoryWebPartProps> {
  private peopleService: PeopleService | undefined;
  private syncService: SyncService | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();

    try {
      // Initialize Microsoft Graph Client
      const graphClient = await this.context.msGraphClientFactory.getClient('3');

      // Initialize PnP SP
      const sp = spfi().using(SPFx(this.context));

      // Initialize Services
      const graphService = new GraphService(graphClient);
      const listService = new ListService(sp);
      this.peopleService = new PeopleService(graphService, listService);
      this.syncService = new SyncService(graphService, listService);

      console.log('People Directory WebPart initialized successfully');
    } catch (error) {
      console.error('Error initializing People Directory WebPart:', error);
      throw error;
    }
  }

  public render(): void {
    if (!this.peopleService || !this.syncService) {
      // Show error state
      this.domElement.innerHTML = `
        <div style="padding: 20px; color: #a4262c; background-color: #fde7e9; border: 1px solid #a4262c; border-radius: 4px;">
          <strong>Error:</strong> Failed to initialize the People Directory. Please refresh the page or contact your administrator.
        </div>
      `;
      return;
    }

    const element: React.ReactElement<IPeopleDirectoryProps> = React.createElement(
      PeopleDirectory,
      {
        peopleService: this.peopleService,
        syncService: this.syncService,
        displayMode: this.displayMode,

        // User Properties
        showEmail: this.properties.showEmail !== false,
        showJobTitle: this.properties.showJobTitle !== false,
        showDepartment: this.properties.showDepartment !== false,
        showOfficeLocation: this.properties.showOfficeLocation !== false,
        showBusinessPhones: this.properties.showBusinessPhones !== false,
        showMobilePhone: this.properties.showMobilePhone !== false,
        showCity: this.properties.showCity !== false,
        showCountry: this.properties.showCountry !== false,
        showCompanyName: this.properties.showCompanyName !== false,
        showEmployeeId: this.properties.showEmployeeId !== false,

        // Font Styling
        fontSize: this.properties.fontSize || 14,
        fontColor: this.properties.fontColor || '#323130',

        // Icon Styling
        iconSize: this.properties.iconSize || 24,
        iconColor: this.properties.iconColor || '#0078d4'
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure which user properties to display and customize styling'
          },
          groups: [
            {
              groupName: 'User Properties to Display',
              groupFields: [
                PropertyPaneToggle('showEmail', {
                  label: 'Show Email',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showJobTitle', {
                  label: 'Show Job Title',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showDepartment', {
                  label: 'Show Department',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showOfficeLocation', {
                  label: 'Show Office Location',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showBusinessPhones', {
                  label: 'Show Business Phones',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showMobilePhone', {
                  label: 'Show Mobile Phone',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showCity', {
                  label: 'Show City',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showCountry', {
                  label: 'Show Country',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showCompanyName', {
                  label: 'Show Company Name',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showEmployeeId', {
                  label: 'Show Employee ID',
                  onText: 'Visible',
                  offText: 'Hidden'
                })
              ]
            },
            {
              groupName: 'Font Styling',
              groupFields: [
                PropertyPaneSlider('fontSize', {
                  label: 'Font Size (px)',
                  min: 10,
                  max: 24,
                  step: 1,
                  value: 14,
                  showValue: true
                }),
                PropertyPaneChoiceGroup('fontColor', {
                  label: 'Font Color',
                  options: [
                    { key: '#323130', text: 'Dark Gray (Default)', iconProps: { officeFabricIconFontName: 'CircleFill' } },
                    { key: '#000000', text: 'Black', iconProps: { officeFabricIconFontName: 'CircleFill' } },
                    { key: '#605E5C', text: 'Medium Gray', iconProps: { officeFabricIconFontName: 'CircleFill' } },
                    { key: '#0078d4', text: 'Blue', iconProps: { officeFabricIconFontName: 'CircleFill' } },
                    { key: '#004578', text: 'Dark Blue', iconProps: { officeFabricIconFontName: 'CircleFill' } }
                  ]
                })
              ]
            },
            {
              groupName: 'Icon Styling',
              groupFields: [
                PropertyPaneSlider('iconSize', {
                  label: 'Icon Size (px)',
                  min: 16,
                  max: 48,
                  step: 2,
                  value: 24,
                  showValue: true
                }),
                PropertyPaneChoiceGroup('iconColor', {
                  label: 'Icon Color',
                  options: [
                    { key: '#0078d4', text: 'Blue (Default)', iconProps: { officeFabricIconFontName: 'CircleFill' } },
                    { key: '#323130', text: 'Dark Gray', iconProps: { officeFabricIconFontName: 'CircleFill' } },
                    { key: '#107c10', text: 'Green', iconProps: { officeFabricIconFontName: 'CircleFill' } },
                    { key: '#8764b8', text: 'Purple', iconProps: { officeFabricIconFontName: 'CircleFill' } },
                    { key: '#d13438', text: 'Red', iconProps: { officeFabricIconFontName: 'CircleFill' } }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
