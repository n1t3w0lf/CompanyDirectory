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
// import * as strings from 'PeopleDirectoryWebPartStrings';
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

        // Header Configuration
        headingText: this.properties.headingText || 'People Directory',
        headingFontSize: this.properties.headingFontSize || 28,
        headingFontColor: this.properties.headingFontColor || '#323130',
        subtextText: this.properties.subtextText || '',
        subtextFontSize: this.properties.subtextFontSize || 14,
        subtextFontColor: this.properties.subtextFontColor || '#666666',

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

        // Profile Name Styling
        profileNameFontSize: this.properties.profileNameFontSize || 18,
        profileNameFontColor: this.properties.profileNameFontColor || '#323130',

        // Profile Properties Styling
        profilePropertiesFontSize: this.properties.profilePropertiesFontSize || 14,
        profilePropertiesFontColor: this.properties.profilePropertiesFontColor || '#605E5C',

        // Text Truncation
        textEllipsisLength: this.properties.textEllipsisLength || 50,

        // Icon Styling
        iconSize: this.properties.iconSize || 20,
        iconColor: this.properties.iconColor || '#0078d4',

        // Profile Card Background
        profileCardBackgroundColor: this.properties.profileCardBackgroundColor || '',
        profileCardBackgroundImage: this.properties.profileCardBackgroundImage || '',

        // Webpart Background
        webpartBackgroundColor: this.properties.webpartBackgroundColor || '',
        webpartBackgroundImage: this.properties.webpartBackgroundImage || ''
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
            description: 'Configure header, styling, and user properties'
          },
          groups: [
            {
              groupName: 'Header Configuration',
              groupFields: [
                PropertyPaneTextField('headingText', {
                  label: 'Heading Text',
                  placeholder: 'People Directory'
                }),
                PropertyPaneSlider('headingFontSize', {
                  label: 'Heading Font Size (px)',
                  min: 20,
                  max: 40,
                  step: 2,
                  value: 28,
                  showValue: true
                }),
                PropertyPaneTextField('headingFontColor', {
                  label: 'Heading Color (hex, rgba, or hsla)',
                  placeholder: '#323130 or rgba(50,49,48,1)',
                  description: 'Examples: #0078d4, rgba(0,120,212,0.8), hsla(206,100%,42%,0.8)'
                }),
                PropertyPaneTextField('subtextText', {
                  label: 'Subtext',
                  placeholder: 'Optional description text',
                  multiline: true,
                  rows: 2
                }),
                PropertyPaneSlider('subtextFontSize', {
                  label: 'Subtext Font Size (px)',
                  min: 10,
                  max: 20,
                  step: 1,
                  value: 14,
                  showValue: true
                }),
                PropertyPaneTextField('subtextFontColor', {
                  label: 'Subtext Color (hex, rgba, or hsla)',
                  placeholder: '#666666 or rgba(102,102,102,1)',
                  description: 'Examples: #666666, rgba(102,102,102,0.9), hsla(0,0%,40%,0.9)'
                })
              ]
            },
            {
              groupName: 'Profile Name Styling',
              groupFields: [
                PropertyPaneSlider('profileNameFontSize', {
                  label: 'Name Font Size (px)',
                  min: 14,
                  max: 28,
                  step: 1,
                  value: 18,
                  showValue: true
                }),
                PropertyPaneTextField('profileNameFontColor', {
                  label: 'Name Color (hex, rgba, or hsla)',
                  placeholder: '#323130 or rgba(50,49,48,1)',
                  description: 'Examples: #000000, rgba(0,0,0,0.9), hsla(0,0%,0%,0.9)'
                })
              ]
            },
            {
              groupName: 'Profile Properties Styling',
              groupFields: [
                PropertyPaneSlider('profilePropertiesFontSize', {
                  label: 'Properties Font Size (px)',
                  min: 10,
                  max: 20,
                  step: 1,
                  value: 14,
                  showValue: true
                }),
                PropertyPaneTextField('profilePropertiesFontColor', {
                  label: 'Properties Color (hex, rgba, or hsla)',
                  placeholder: '#605E5C or rgba(96,94,92,1)',
                  description: 'Examples: #605E5C, rgba(96,94,92,0.9), hsla(20,3%,37%,0.9)'
                })
              ]
            },
            {
              groupName: 'Text Truncation',
              groupFields: [
                PropertyPaneSlider('textEllipsisLength', {
                  label: 'Maximum Characters Before Ellipsis',
                  min: 20,
                  max: 200,
                  step: 10,
                  value: 50,
                  showValue: true
                })
              ]
            },
            {
              groupName: 'Icon Styling',
              groupFields: [
                PropertyPaneSlider('iconSize', {
                  label: 'Icon Size (px)',
                  min: 16,
                  max: 32,
                  step: 2,
                  value: 20,
                  showValue: true
                }),
                PropertyPaneTextField('iconColor', {
                  label: 'Icon Color (hex, rgba, or hsla)',
                  placeholder: '#0078d4 or rgba(0,120,212,1)',
                  description: 'Examples: #0078d4, rgba(0,120,212,0.8), hsla(206,100%,42%,0.8)'
                })
              ]
            },
            {
              groupName: 'Profile Card Background',
              groupFields: [
                PropertyPaneTextField('profileCardBackgroundColor', {
                  label: 'Background Color (hex, rgba, or hsla)',
                  placeholder: '#ffffff or rgba(255,255,255,1)',
                  description: 'Examples: #f3f2f1, rgba(243,242,241,0.9), hsla(0,0%,95%,0.9)'
                }),
                PropertyPaneTextField('profileCardBackgroundImage', {
                  label: 'Background Image URL (optional)',
                  placeholder: 'https://example.com/image.jpg',
                  description: 'Leave empty to use color only. Image will override color if both are set.'
                })
              ]
            },
            {
              groupName: 'Webpart Background',
              groupFields: [
                PropertyPaneTextField('webpartBackgroundColor', {
                  label: 'Background Color (hex, rgba, or hsla)',
                  placeholder: '#ffffff or rgba(255,255,255,1)',
                  description: 'Examples: #faf9f8, rgba(250,249,248,1), hsla(30,20%,97%,1)'
                }),
                PropertyPaneTextField('webpartBackgroundImage', {
                  label: 'Background Image URL (optional)',
                  placeholder: 'https://example.com/background.jpg',
                  description: 'Leave empty to use color only. Image will override color if both are set.'
                })
              ]
            },
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
            }
          ]
        }
      ]
    };
  }
}
