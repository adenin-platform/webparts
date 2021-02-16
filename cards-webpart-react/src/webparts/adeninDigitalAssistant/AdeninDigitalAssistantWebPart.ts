import * as React from 'react';
import * as ReactDom from 'react-dom';
import ReactHtmlParser from 'react-html-parser'; 
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldTextWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldTextWithCallout';
import { PropertyFieldDropdownWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldDropdownWithCallout';
import { PropertyFieldMessage} from '@pnp/spfx-property-controls/lib/PropertyFieldMessage';
import { MessageBarType } from 'office-ui-fabric-react';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {EmptyControl} from './components/PropertyControls';

import * as strings from 'AdeninDigitalAssistantWebPartStrings';
import AdeninDigitalAssistant from './components/AdeninDigitalAssistant';
import { IAdeninDigitalAssistantProps } from './components/IAdeninDigitalAssistantProps';
import UserHelp from './components/UserHelp';

// Default settings
const adeninTenantId: string = 'ce4cc661-4506-4d48-8c64-c5b090aa46fb';

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

  private validateURL(url: string): string {
    if (url === null ||
      url.trim().length === 0 ||
      !(/^(https?|ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(\#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i.test(url.trim()))) {
      return 'Please enter a valid URL';
    }
    return '';
  }

  private validateClassList(classes: string): string {
    if (classes != null && classes.trim().length > 0 && !/^([a-z_]|-[a-z_-])[a-z\d_-]*$/i.test(classes)) {
      return 'Please enter a valid class name';
    }
    return '';
  }

  private validateUUID(uuid: string): string {
    if (uuid != null && uuid.trim().length > 0 && !/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(uuid)) {
      return 'Please enter a valid Card ID';
    }
    return '';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let cardIdTextbox = (this.properties.embedType == 'card') ? 
                        PropertyPaneTextField('cardId', {
                          label: "Card Id",
                          onGetErrorMessage: this.validateUUID.bind(this)
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
    
    let componentCDNTextbox = ((this.context.pageContext.aadInfo ? this.context.pageContext.aadInfo.tenantId._guid : '') == adeninTenantId) ? 
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
                PropertyFieldTextWithCallout('tenantURL', {
                  calloutTrigger: CalloutTriggers.Hover,
                  key: 'tenantURL',
                  label: strings.tenantURLLabel,
                  value: this.properties.tenantURL,
                  calloutContent: React.createElement('span', {}, strings.tenantURLDescription),
                  calloutWidth: 200,
                  onGetErrorMessage: this.validateURL.bind(this)
                }),
                PropertyFieldDropdownWithCallout('embedType', {
                  calloutTrigger: CalloutTriggers.Hover,
                  key: 'embedType',
                  label: strings.embedTypeDropdownLabel,
                  options: [
                    { key: 'searchCard', text: 'Search results'},
                    { key: 'card', text: 'Card'},
                    { key: 'board', text: 'Board'}
                  ],
                  selectedKey: this.properties.embedType,
                  calloutContent: this.embedTypeCalloutContent(),
                  calloutWidth: 265,
                }),
                PropertyFieldMessage("", {
                  key: "cardEmbedMessage",
                  text: ReactHtmlParser(strings.cardEmbedMessage),
                  messageType:  MessageBarType.info,
                  isVisible: (this.properties.embedType == 'card'),
                }),
                PropertyFieldMessage("", {
                  key: "searchEmbedMessage",
                  text: ReactHtmlParser(strings.searchEmbedMessage),
                  messageType:  MessageBarType.info,
                  isVisible:  (this.properties.embedType == 'searchCard'),
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
                PropertyFieldTextWithCallout('customCSSClasses', {
                  calloutTrigger: CalloutTriggers.Hover,
                  key: 'customCSSClasses',
                  label: strings.customCSSLabel,
                  description: strings.customCSSDescription,
                  value: this.properties.customCSSClasses,
                  calloutContent: React.createElement('span', {}, 'Optionally enter a list of classes to apply to the root Card element'),
                  calloutWidth: 200,
                  onGetErrorMessage: this.validateClassList.bind(this)
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

  private embedTypeCalloutContent(): JSX.Element {
    return React.createElement('div', {}, ReactHtmlParser('<strong>Search results</strong><br/>Show Assistant Cards alongside search results on custom search pages<br/><br/><strong>Card</strong><br/>Show a single Assistant Card<br/><br/><strong>Board</strong><br/>Show the current user\'s Board'));
  }
}

