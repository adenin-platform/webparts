import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {EmptyControl} from './components/PropertyControls';

import * as strings from 'AdeninDigitalAssistantWebPartStrings';
import AdeninDigitalAssistant from './components/AdeninDigitalAssistant';
import { IAdeninDigitalAssistantProps } from './components/IAdeninDigitalAssistantProps';
import UserHelp from './components/UserHelp';

// Default settings
const defaultCDN: string = 'https://components.adenin.com/components';
const defaultClientID: string = 'c44ce7b8-8f45-4ec6-9f7e-e4a80f9d8edc';
const defaultSSOProviderID: string = '';
const defaultContextLoader: string = '/at-app/at-app-context-oidc.js';
const toasterTenantId: string = 'ce4cc661-4506-4d48-8c64-c5b090aa46fb';

export default class AdeninDigitalAssistantWebPart extends BaseClientSideWebPart<IAdeninDigitalAssistantProps> {

  public render(): void {
    const element: React.ReactElement<IAdeninDigitalAssistantProps> = React.createElement(
      AdeninDigitalAssistant,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        context: this.context,
        tenantId: this.context.pageContext.aadInfo ? this.context.pageContext.aadInfo.tenantId._guid : '',
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        tenantURL: this.properties.tenantURL,
        componentCDN: this.properties.componentCDN,
        SSOProviderID: this.properties.SSOProviderID,
        componentClientID: this.properties.componentClientID,
        embedType: this.properties.embedType,
        cardId: this.properties.cardId,
        cardStyle: this.properties.cardStyle,
        customCSSClasses: this.properties.customCSSClasses,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let cardIdTextbox = (this.properties.embedType == 'card') ? 
                        PropertyPaneTextField('cardId', {
                          label: "Card Id"
                        }) :
                        this.emptyControl;

    let cardStyleDropdown = (this.properties.embedType == 'card') ? 
                            PropertyPaneDropdown('cardStyle', {
                              label: "Card Style",
                              options: [
                                { key: 'none', text: 'No box'},
                                { key: 'card', text: 'Card'}
                              ],
                              selectedKey: 'card',
                            }) :
                            this.emptyControl;
    
    let componentCDNTextbox = (this.properties.tenantId == toasterTenantId) ? 
                              PropertyPaneTextField('componentCDN', {
                                label: strings.componentCDNFieldLabel
                              }) : 
                              this.emptyControl;

    return {
      pages: [
        {
          header: {
            description: strings.propertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.BasicGroupName,
              isCollapsed: false,
              groupFields: [
                PropertyPaneLabel('label', {
                  text: strings.tenantURLLabel,
                  required: true,
                }),
                PropertyPaneTextField('tenantURL', {
                  description: strings.tenantURLDescription
                }),
                PropertyPaneLabel('label', {
                  text: strings.embedTypeDropdownLabel,
                  required: true
                }),
                PropertyPaneDropdown('embedType', {
                  label: null,
                  options: [
                    { key: 'searchCard', text: 'Search results'},
                    { key: 'card', text: 'Card'},
                    { key: 'board', text: 'Board'}
                  ]
                }),
                cardIdTextbox,
                cardStyleDropdown,
                new UserHelp()
              ]
            },
            {
              groupName: strings.componentConfigName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneLabel('label', {
                  text: strings.leaveBlankForDefault,
                }),
                PropertyPaneTextField('customCSSClasses', {
                  label: strings.customCSSLabel,
                  description: strings.customCSSDescription
                }),
                PropertyPaneTextField('SSOProviderID', {
                  label: strings.componentSSOProviderIDFieldLabel
                }),
                PropertyPaneTextField('componentClientID', {
                  label: strings.componentClientIDFieldLabel
                }),
                componentCDNTextbox,
              ],
            },
          ]
        }
      ]
    };
  }

  // Implement Empty Control class for hidden IPropertyPaneFields
  private _emptyControl : EmptyControl = null;

  public get emptyControl(): EmptyControl {
    if (this._emptyControl == null) {
      this._emptyControl  = new EmptyControl();
    }
    return this._emptyControl;
  }
}

