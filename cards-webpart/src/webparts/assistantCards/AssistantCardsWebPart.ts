import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AssistantCardsWebPart.module.scss';
import * as strings from 'AssistantCardsWebPartStrings';

import {EmptyControl} from './components/PropertyControls';

export interface IAssistantCardsWebPartProps {
  APIEndpoint: string;
  embedType: string;
  cardId: string;
  cardStyle: string;
}

export default class AssistantCardsWebPart extends BaseClientSideWebPart<IAssistantCardsWebPartProps> {

  public renderConfigMessage() {
    let message = this.properties.APIEndpoint ? 
    `<p class="${ styles.url}">Tenant URL: ${escape(this.properties.APIEndpoint)}</p>
    <p class="${ styles.url}">Display mode: ${this.properties.embedType ? escape(this.properties.embedType) : "Not set"}</p>`
    :
    `<p class="${ styles.url}">Please configure your tenant URL in the webpart settings.</p>`;

    return (`
        <div class="${ styles.assistantCards}">
          <div class="${ styles.container}">
            <div class="${ styles.row}">
              <div class="${ styles.column}">
                <span class="${ styles.title}">Digital Assistant Cards Webpart</span>
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
        <div class="${ styles.wrapper}">
          <at-app-card name='${this.properties.cardId}' card-container-type='modal' box='${this.properties.cardStyle}' push></at-app-card>
        </div>`
      );
    } else if (this.properties.embedType == "searchCard") {
      return (`
        <div class="${ styles.wrapper}">
          <include-element id="intent-card" name="at-intent-card/at-intent-card.html" event-source-selector=".ms-SearchBox-field"></include-element>
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
      
      if (this.properties.APIEndpoint) {
        this.loadScript(this.properties.APIEndpoint);
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
                            label: "Card ID"
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
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('APIEndpoint', {
                  label: strings.APIEndpointFieldLabel
                })
              ]
            },
            {
              groupName: strings.embedTypeName,
              groupFields: [
                PropertyPaneDropdown('embedType', {
                  label: strings.embedTypeDropdownLabel,
                  options: [
                    { key: 'searchCard', text: 'Search Result Card'},
                    { key: 'card', text: 'Card'}
                  ]
                }),
                cardIdTextbox,
                cardStyleDropdown,
              ]
            }
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
  private loadScript(endpoint:string) {
    // Trim trailing slash from endpoint if present
    endpoint = endpoint.replace(/\/$/, "");

    var contextLoader = () => {
      var script = document.createElement('script');
      script.src = 'https://components.adenin.com/components/at-app/at-app-context-es6.js';
      script.id='at-app-context';
      script.setAttribute('session-service-url', endpoint+'/session/myprofile');
      script.async = true;
      document.head.appendChild(script);
    };

    if(document.readyState=="complete" ) {
      contextLoader();
    } else {
      window.onload = contextLoader;
    }
  }
}
