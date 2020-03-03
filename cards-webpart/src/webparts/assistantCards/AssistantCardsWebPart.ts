import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AssistantCardsWebPart.module.scss';
import * as strings from 'AssistantCardsWebPartStrings';

export interface IAssistantCardsWebPartProps {
  APIEndpoint: string;
}

export default class AssistantCardsWebPart extends BaseClientSideWebPart<IAssistantCardsWebPartProps> {

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
        this.domElement.innerHTML = `
        <script>
          var contextLoader = function(e) {
              var script = document.createElement('script');
              script.src = 'https://components.adenin.com/components/at-app/at-app-context-es6.js';
              script.id='at-app-context';
              script.setAttribute('session-service-url', '${ this.properties.APIEndpoint}/session/myprofile');
              script.async = true;
              document.head.appendChild(script);
          };
  
          if(document.readyState=="complete" ) {
          contextLoader();
          } else {
          window.onload =contextLoader;
          }
        </script>
  
        <div class="${ styles.wrapper}">
          <include-element id="intent-card" name="at-intent-card/at-intent-card.html" event-source-selector=".ms-SearchBox-field"></include-element>
        </div>`;
        this.executeScript(this.domElement);
      } else {
        this.domElement.innerHTML = `
        <div class="${ styles.assistantCards}">
          <div class="${ styles.container}">
            <div class="${ styles.row}">
              <div class="${ styles.column}">
                <span class="${ styles.title}">Digital Assistant Cards Webpart</span>
                <p class="${ styles.url}">Please configure your tenant URL in the webpart settings.</p>
              </div>
            </div>
          </div>
        </div>`;
      }

    } else {
      this.domElement.innerHTML = `
          <div class="${ styles.assistantCards}">
            <div class="${ styles.container}">
              <div class="${ styles.row}">
                <div class="${ styles.column}">
                  <span class="${ styles.title}">Digital Assistant Cards Webpart</span>
                  <p class="${ styles.url}">Tenant URL: ${escape(this.properties.APIEndpoint)}</p>
                </div>
              </div>
            </div>
          </div>`;
    }

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
                PropertyPaneTextField('APIEndpoint', {
                  label: strings.APIEndpointFieldLabel
                })
              ]
            },
          ]
        }
      ]
    };
  }

  private evalScript(elem) {
    const data = (elem.text || elem.textContent || elem.innerHTML || "");
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    const scriptTag = document.createElement("script");

    scriptTag.type = "text/javascript";
    if (elem.src && elem.src.length > 0) {
      return;
    }
    if (elem.onload && elem.onload.length > 0) {
      scriptTag.onload = elem.onload;
    }

    try {
      // doesn't work on ie...
      scriptTag.appendChild(document.createTextNode(data));
    } catch (e) {
      // IE has funky script nodes
      scriptTag.text = data;
    }

    headTag.insertBefore(scriptTag, headTag.firstChild);
    headTag.removeChild(scriptTag);
  }

  private nodeName(elem, name) {
    return elem.nodeName && elem.nodeName.toUpperCase() === name.toUpperCase();
  }

  // Finds and executes scripts in a newly added element's body.
  // Needed since innerHTML does not run scripts.
  //
  // Argument element is an element in the dom.
  private async executeScript(element: HTMLElement) {
    // Define global name to tack scripts on in case script to be loaded is not AMD/UMD

    (<any>window).ScriptGlobal = {};

    // main section of function
    const scripts = [];
    const children_nodes = element.childNodes;

    for (let i = 0; children_nodes[i]; i++) {
      const child: any = children_nodes[i];
      if (this.nodeName(child, "script") &&
        (!child.type || child.type.toLowerCase() === "text/javascript")) {
        scripts.push(child);
      }
    }

    const urls = [];
    const onLoads = [];
    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.src && scriptTag.src.length > 0) {
        urls.push(scriptTag.src);
      }
      if (scriptTag.onload && scriptTag.onload.length > 0) {
        onLoads.push(scriptTag.onload);
      }
    }

    let oldamd = null;
    if (window["define"] && window["define"].amd) {
      oldamd = window["define"].amd;
      window["define"].amd = null;
    }


    for (let i = 0; i < urls.length; i++) {
      try {
        let scriptUrl = urls[i];
        const prefix = scriptUrl.indexOf('?') === -1 ? '?' : '&';
        scriptUrl += prefix + 'cow=' + new Date().getTime();
        await SPComponentLoader.loadScript(scriptUrl, { globalExportsName: "ScriptGlobal" });
      } catch (error) {
        if (console.error) {
          console.error(error);
        }
      }
    }
    if (oldamd) {
      window["define"].amd = oldamd;
    }

    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.parentNode) { scriptTag.parentNode.removeChild(scriptTag); }
      this.evalScript(scripts[i]);
    }
    // execute any onload people have added
    for (let i = 0; onLoads[i]; i++) {
      onLoads[i]();
    }
  }
}
