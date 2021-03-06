import * as React from 'react';
import { IAdeninDigitalAssistantProps } from './IAdeninDigitalAssistantProps';
import ErrorBoundary from "./ErrorBoundary";
import * as AdaptiveCards from "adaptivecards";
import * as ACData from "adaptivecards-templating";
import * as markdownit from "markdown-it";

const defaultSearchCSSClasses: string = 'loading sp_search';
const defaultCardCSSClasses: string = 'loading webpart';
const defaultCDN: string = 'https://components.adenin.com/components';
const defaultClientID: string = 'c44ce7b8-8f45-4ec6-9f7e-e4a80f9d8edc';
const defaultSSOProviderID: string = '';

declare global {
    namespace JSX {
        interface IntrinsicElements {
            'at-app-card': React.DetailedHTMLProps<React.HTMLAttributes<HTMLElement>, HTMLElement>;
            'include-element': React.DetailedHTMLProps<React.HTMLAttributes<HTMLElement>, HTMLElement>;
        }
    }
}

declare module "react" {
    interface HTMLAttributes<T> {
        name?: string;
        box?: string;
        class?: string;
        indicator?: boolean;
    }
}

export class Card extends React.Component<IAdeninDigitalAssistantProps> {

    public state = {
        searchBoxVisible: false,
    };

    constructor(props) {
        super(props);

        // set window variables
        console.log('setting window variables');

        var testingEnvAPI = (localStorage.getItem("testingEnvAPI") && localStorage.getItem("testingEnvAPI").length > 0) ? localStorage.getItem("testingEnvAPI") : null;

        window["Tangere"] = window["Tangere"] || {};
        window["AdaptiveCards"] = AdaptiveCards;
        window["ACData"] = ACData;
        window["markdownit"] = markdownit;

        if (!window["Tangere"].identity) {
            window["Tangere"].identity = {
                session_service_url: (testingEnvAPI ? testingEnvAPI.trim().replace(/\/+$/, "") : this.props.tenantURL.trim().replace(/\/+$/, "")) + '/session/myprofile',
                provider_id: this.props.SSOProviderID ? this.props.SSOProviderID : defaultSSOProviderID,
                client_id: this.props.componentClientID ? this.props.componentClientID : defaultClientID,
                redirect_uri: (this.props.componentCDN ? this.props.componentCDN.trim().replace(/\/+$/, "") : defaultCDN) + "/sso/passiveCallback.html",
                authorization: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
                login_hint: this.props.context.pageContext.user.loginName,
                token_issuer: "aad." + (this.props.context.pageContext.aadInfo ? this.props.context.pageContext.aadInfo.tenantId._guid : '')
            };
        }
    }

    public componentDidMount() {
        var observer = new MutationObserver((mutations, me) => {
            var searchbox = document.querySelector('.ms-SearchBox-field');
            if (searchbox) {
                console.log('found the searchbox');
                me.disconnect();
                this.setState({searchBoxVisible: true});
            }
        });
          
        observer.observe(document, {
            childList: true,
            subtree: true
        });

        var loadScript = (src:string) => {
          var tag = document.createElement('script');
          tag.async = false;
          tag.src = src;
          var body = document.getElementsByTagName('body')[0];
          body.appendChild(tag);
        };
    
        try {
            console.log('loading component script');
            loadScript((this.props.componentCDN ? this.props.componentCDN.trim().replace(/\/+$/, "") : defaultCDN) + '/at-app/at-app-context-oidc.js');
        } catch (error) {
          throw new Error(error);
        }
    }

    public render(): React.ReactElement<IAdeninDigitalAssistantProps> {
        if (this.props.embedType == "card") {
            return (
                <ErrorBoundary>
                <div>
                    <at-app-card class={this.props.customCSSClasses ? this.props.customCSSClasses : defaultCardCSSClasses} style={{maxWidth: 450 + "px"}} name={this.props.cardId} card-container-type="sp-card" box={this.props.cardStyle}></at-app-card>
                </div>
                </ErrorBoundary>
            );
        } else if (this.props.embedType == "searchCard" && this.state.searchBoxVisible) {
            return (
                <ErrorBoundary>
                <div>
                    <include-element id="intent-card" name="at-intent-card/at-intent-card.html" class={this.props.customCSSClasses ? this.props.customCSSClasses : defaultSearchCSSClasses} card-container-type='sp-search' event-source-selector=".ms-SearchBox-field" indicator></include-element>
                </div>
                </ErrorBoundary>
            );
        } else if (this.props.embedType == "board") {
            return (
                <ErrorBoundary>
                <div>
                    <include-element id="board" name="at-card-board/at-card-board.html" class={this.props.customCSSClasses ? this.props.customCSSClasses : defaultCardCSSClasses}></include-element>
                </div>
                </ErrorBoundary>
            );
        } else {
            return null;
        }
    }
}