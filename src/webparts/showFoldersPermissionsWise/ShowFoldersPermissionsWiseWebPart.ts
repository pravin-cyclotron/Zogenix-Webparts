import * as React from 'react';
import "@pnp/polyfill-ie11";
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel,
  PropertyPaneLink
} from '@microsoft/sp-webpart-base';

import * as strings from 'ShowFoldersPermissionsWiseWebPartStrings';
import ShowFoldersPermissionsWise from './components/ShowFoldersPermissionsWise';
import { IShowFoldersPermissionsWiseProps } from './components/IShowFoldersPermissionsWiseProps';
import { sp } from "@pnp/sp/presets/all";
import { stringIsNullOrEmpty } from '@pnp/common';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from "@pnp/spfx-property-controls/lib/PropertyFieldColorPicker";

export interface IShowFoldersPermissionsWiseWebPartProps {
  siteURL: string;
  title: string;
  iconName: string;
  noFoldersFoundMessage: string;
  titleColor: string;
  backgroundColor: string;
  iconBackgroundColor: string;
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
        iconName: this.properties.iconName,
        noFoldersFoundMessage: this.properties.noFoldersFoundMessage,
        title: this.properties.title,
        displayMode: this.displayMode,
        titleColor: this.properties.titleColor,
        backgroundColor: this.properties.backgroundColor,
        iconBackgroundColor: this.properties.iconBackgroundColor,
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
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('siteURL', {
                  label: strings.SiteURLFieldLabel
                }),
                PropertyPaneTextField('iconName', {
                  label: strings.IconNameFieldLabel
                }),
                PropertyPaneLabel('iconNameDescription', {
                  text: strings.IconDescriptionFieldLabel
                }),
                PropertyPaneLink('iconSourceLink', {
                  text: "Available Icons List",
                  href: "https://developer.microsoft.com/en-us/fluentui#/styles/web/icons#available-icons",
                  target: "_blank"
                }),
                PropertyPaneTextField('noFoldersFoundMessage', {
                  label: strings.NoFoldersFoundMessage
                }),

                PropertyFieldColorPicker("titleColor", {
                  label: "Text color",
                  selectedColor: this.properties.titleColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: "Precipitation",
                  key: "colorFieldId"
                }),
                PropertyFieldColorPicker("iconBackgroundColor", {
                  label: "Icon Color",
                  selectedColor: this.properties.iconBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: "Precipitation",
                  key: "colorFieldId"
                }),
                PropertyFieldColorPicker("backgroundColor", {
                  label: "Background Color",
                  selectedColor: this.properties.backgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: "Precipitation",
                  key: "colorFieldId"
                })

              ]
            }
          ]
        }
      ]
    };
  }
}
