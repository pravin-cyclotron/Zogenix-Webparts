import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ShowFoldersPermissionsWiseWebPartStrings';
import ShowFoldersPermissionsWise from './components/ShowFoldersPermissionsWise';
import { IShowFoldersPermissionsWiseProps } from './components/IShowFoldersPermissionsWiseProps';
import { sp } from "@pnp/sp/presets/all";
import { stringIsNullOrEmpty } from '@pnp/common';


export interface IShowFoldersPermissionsWiseWebPartProps {
  siteURL: string;
  title: string;
}

export default class ShowFoldersPermissionsWiseWebPart extends BaseClientSideWebPart<IShowFoldersPermissionsWiseWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {

      sp.setup({
        sp: {
          headers: {
            Accept: "application/json;odata=verbose",
          },
          baseUrl: !stringIsNullOrEmpty(this.properties.siteURL) ? this.properties.siteURL : this.context.pageContext.web.absoluteUrl,
        },
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IShowFoldersPermissionsWiseProps> = React.createElement(
      ShowFoldersPermissionsWise,
      {
        siteURL: !stringIsNullOrEmpty(this.properties.siteURL) ? this.properties.siteURL : this.context.pageContext.web.absoluteUrl,
        title: this.properties.title,
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
                PropertyPaneTextField('siteURL', {
                  label: strings.SiteURLFieldLabel
                }),
                PropertyPaneTextField('siteURL', {
                  label: strings.SiteURLFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
