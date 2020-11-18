import * as React from 'react';
import { IAdeninDigitalAssistantProps } from './IAdeninDigitalAssistantProps';

const defaultSearchCSSClasses: string = 'loading sp_search';
const defaultCardCSSClasses: string = 'loading webpart';

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
    public render(): React.ReactElement<IAdeninDigitalAssistantProps> {
        if (this.props.embedType == "card") {
            return (
                <div>
                    <at-app-card class={this.props.customCSSClasses ? this.props.customCSSClasses : defaultCardCSSClasses} name={this.props.cardId} card-container-type="sp-card" box={this.props.cardStyle}></at-app-card>
                </div>
            );
        } else if (this.props.embedType == "searchCard") {
            return (
                <div>
                    <include-element id="intent-card" name="at-intent-card/at-intent-card.html" class={this.props.customCSSClasses ? this.props.customCSSClasses : defaultSearchCSSClasses} card-container-type='sp-search' event-source-selector=".ms-SearchBox-field" indicator></include-element>
                </div>
            );
        } else if (this.props.embedType == "board") {
            return (
                <div>
                    <include-element id="board" name="at-card-board/at-card-board.html" class={this.props.customCSSClasses ? this.props.customCSSClasses : defaultCardCSSClasses}></include-element>
                </div>
            );
        } else {
            return null;
        }
    }
}