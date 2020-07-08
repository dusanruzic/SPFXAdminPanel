import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'AdminPanelWebPartStrings';
import AdminPanel from './components/AdminPanel';
import { IAdminPanelProps } from './components/IAdminPanelProps';
import SharePointService from '../../services/SharePoint/SharePointService';
import {Environment } from '@microsoft/sp-core-library';

export interface IAdminPanelWebPartProps {
  description: string;
}

export default class AdminPanelWebPart extends BaseClientSideWebPart<IAdminPanelWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAdminPanelProps > = React.createElement(
      AdminPanel,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then(() =>{

      SharePointService.setup(this.context, Environment.type);

    });
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
