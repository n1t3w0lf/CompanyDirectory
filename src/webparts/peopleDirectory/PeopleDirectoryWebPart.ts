import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'PeopleDirectoryWebPartStrings';
import { PeopleDirectory } from './components/PeopleDirectory';
import { IPeopleDirectoryProps } from './components/IPeopleDirectoryProps';
import { GraphService } from '../../services/GraphService';
import { ListService } from '../../services/ListService';
import { PeopleService } from '../../services/PeopleService';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export interface IPeopleDirectoryWebPartProps {
  title: string;
  description: string;
}

export default class PeopleDirectoryWebPart extends BaseClientSideWebPart<IPeopleDirectoryWebPartProps> {
  private peopleService: PeopleService | undefined;

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

      console.log('People Directory WebPart initialized successfully');
    } catch (error) {
      console.error('Error initializing People Directory WebPart:', error);
      throw error;
    }
  }

  public render(): void {
    if (!this.peopleService) {
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
        title: this.properties.title || 'People Directory',
        description: this.properties.description || 'Search and discover people across your organization',
        peopleService: this.peopleService,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  description: 'The title displayed at the top of the web part'
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  description: 'A brief description shown below the title',
                  multiline: true,
                  rows: 3
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
