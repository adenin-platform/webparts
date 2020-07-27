import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AssistantCardsWebPart.module.scss';
import * as strings from 'AssistantCardsWebPartStrings';

// Default settings
const defaultCDN: string = 'https://components.adenin.com/components';
const defaultClientID: string = 'c44ce7b8-8f45-4ec6-9f7e-e4a80f9d8edc';
const defaultSSOProviderID: string = '';
const defaultContextLoader: string = '/at-app/at-app-context-oidc.js';
const defaultSearchCSSClasses: string = 'loading sp_search';
const defaultCardCSSClasses: string = 'loading webpart';

import {EmptyControl} from './components/PropertyControls';

export interface IAssistantCardsWebPartProps {
  tenantURL: string;
  componentCDN: string;
  SSOProviderID: string;
  componentClientID: string;
  contextLoaderSrc: string;
  embedType: string;
  cardId: string;
  cardStyle: string;
  customCSSClasses: string;
}

export default class AssistantCardsWebPart extends BaseClientSideWebPart<IAssistantCardsWebPartProps> {

  public renderConfigMessage() {
    let message = this.properties.tenantURL ? 
      `<p class="${ styles.url}">Tenant URL: ${escape(this.properties.tenantURL)}</p>
      <p class="${ styles.url}">Display mode: ${this.properties.embedType ? escape(this.properties.embedType) : "Not set"}</p>
      ${this.properties.componentCDN ? `<p class="${ styles.url}">Component CDN URL: ${this.properties.componentCDN}</p>` : ''}`
      :
      `<p class="${ styles.url}">Please configure your tenant URL in the web part settings.</p>`;

    return (`
        <div class="${ styles.assistantCards}">
          <div class="${ styles.container}">
            <div class="${ styles.row}">
              <div class="${ styles.column}">
                <span class="${ styles.title}">Digital Assistant Cards web part</span>
                ${message}
              </div>
            </div>
          </div>
        </div>`
    );
  }

  public renderCard() {
    if (this.properties.embedType == "card") {
      return (`
        <div>
          <at-app-card class="${this.properties.customCSSClasses ? this.properties.customCSSClasses : defaultCardCSSClasses}" name="${this.properties.cardId}" card-container-type="sp-card" box="${this.properties.cardStyle}"></at-app-card>
        </div>`
      );
    } else if (this.properties.embedType == "searchCard") {
      return (`
        <div>
          <include-element id="intent-card" name="at-intent-card/at-intent-card.html" class="${this.properties.customCSSClasses ? this.properties.customCSSClasses : defaultSearchCSSClasses}" card-container-type='sp-search' event-source-selector=".ms-SearchBox-field" indicator></include-element>
        </div>`
      );
    } else if (this.properties.embedType == "board") {
      return (`
        <div>
          <include-element id="board" name="at-card-board/at-card-board.html" class="${this.properties.customCSSClasses ? this.properties.customCSSClasses : defaultCardCSSClasses}"></include-element>
        </div>`
      );
    } else {
      return null;
    }
  }

  public render(): void {

    if (this.displayMode == DisplayMode.Read) {
      let element = this.domElement.parentElement;
      // check up to 5 levels up for padding and exit once found
      for (let i = 0; i < 5; i++) {
        const style = window.getComputedStyle(element);
        const hasPadding = style.paddingTop !== "0px";
        if (hasPadding) {
          element.style.paddingTop = "0px";
          element.style.paddingBottom = "0px";
          element.style.marginTop = "0px";
          element.style.marginBottom = "0px";
        }
        element = element.parentElement;
      }
      
      if (this.properties.tenantURL) {
        this.loadScript(this.properties.tenantURL, this.properties.componentCDN, this.properties.SSOProviderID, this.properties.componentClientID, this.properties.contextLoaderSrc);
        this.domElement.innerHTML = this.renderCard();
      } else {
        this.domElement.innerHTML = this.renderConfigMessage();
      }
    } else {
      this.domElement.innerHTML = this.renderConfigMessage();
    }
  }



  protected get dataVersion(): Version {
    return Version.parse('1.0');
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

    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneLabel('label', {
                  text: strings.tenantURLLabel,
                  required: true
                }),
                PropertyPaneTextField('tenantURL', {}),
                PropertyPaneLabel('label', {
                  text: strings.embedTypeDropdownLabel,
                  required: true
                }),
                PropertyPaneDropdown('embedType', {
                  label: null,
                  options: [
                    { key: 'searchCard', text: 'Search Result Card'},
                    { key: 'card', text: 'Card'},
                    { key: 'board', text: 'Board'}
                  ]
                }),
                cardIdTextbox,
                cardStyleDropdown,
              ]
            },
            {
              groupName: strings.componentConfigName,
              groupFields: [
                PropertyPaneLabel('label', {
                  text: strings.leaveBlankForDefault,
                }),
                PropertyPaneTextField('componentCDN', {
                  label: strings.componentCDNFieldLabel
                }),
                PropertyPaneTextField('SSOProviderID', {
                  label: strings.componentSSOProviderIDFieldLabel
                }),
                PropertyPaneTextField('componentClientID', {
                  label: strings.componentClientIDFieldLabel
                }),
                PropertyPaneTextField('contextLoaderSrc', {
                  label: strings.contextLoaderLabel
                }),
                PropertyPaneTextField('customCSSClasses', {
                  label: strings.customCSSLabel,
                  description: strings.customCSSDescription
                })
              ],
              isCollapsed: true,
            },
          ]
        },
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

  // Load adenin Card components
  private loadScript(tenantURL:string, componentCDN: string, SSOProviderID: string, componentClientID: string, contextLoaderSrc: string) {
    // Trim trailing slash from CDN if present
    let cdnURL = componentCDN ? componentCDN.replace(/\/$/, "") : defaultCDN;
    let endpoint = tenantURL ? tenantURL.replace(/\/$/, "") + '/session/myprofile' : null;
    let contextLoaderScript = contextLoaderSrc ? contextLoaderSrc.trim() :  defaultContextLoader;

    var contextLoader = () => {
      console.log("adenin webpart");

      window["Tangere"] = window["Tangere"] || {};
      window["Tangere"].identity = {
        session_service_url: endpoint,
        provider_id: SSOProviderID ? SSOProviderID : defaultSSOProviderID,
        client_id: componentClientID ? componentClientID : defaultClientID,
        redirect_uri: cdnURL + "/sso/passiveCallback.html",
        authorization: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
        login_hint: window["_spPageContextInfo"].userPrincipalName,
        token_issuer: "aad." + window["_spPageContextInfo"].aadTenantId
      };

      var script = document.createElement('script');
      script.src = cdnURL + contextLoaderScript;
      script.async = true;
      document.head.appendChild(script);
    };

    window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;

    if(document.readyState=="complete" ) {
      contextLoader();
    } else {
      window.onload = contextLoader;
    }
  }
}
