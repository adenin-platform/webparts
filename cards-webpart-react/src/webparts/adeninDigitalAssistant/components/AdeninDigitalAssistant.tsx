import * as React from 'react';
import styles from './AdeninDigitalAssistant.module.scss';
import { IAdeninDigitalAssistantProps } from './IAdeninDigitalAssistantProps';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Text } from 'office-ui-fabric-react/lib/Text';
import { Placeholder } from "./Placeholder/Placeholder";
import { Card } from "./Card";
import ErrorBoundary from "./ErrorBoundary";
import { DisplayMode } from '@microsoft/sp-core-library';
import * as strings from 'AdeninDigitalAssistantWebPartStrings';

export default class AdeninDigitalAssistant extends React.Component<IAdeninDigitalAssistantProps, {}> {

  public render(): React.ReactElement<IAdeninDigitalAssistantProps> {
    let renderPlaceholder: JSX.Element = null;
    let renderCard: JSX.Element = null;

    if (this.props.displayMode === DisplayMode.Edit) {
      if ((!this.props.tenantURL || !this.props.embedType)) {
        renderPlaceholder = <Placeholder iconName='digital-assistant'
                                     iconText='adenin Digital Assistant'
                                     contentClassName={styles.placeholder}
                                     description={strings.ShowBlankEditMessage}
                                     configButtonLabel='Configure'
                                     assistantButtonLabel='Enable in Digital Assistant'
                                     onConfigure={this._onConfigure} />;
      } else {
        renderPlaceholder = <Placeholder iconName='digital-assistant'
                                     iconText='adenin Digital Assistant'
                                     contentClassName={styles.placeholder}
                                     description={strings.ShowBlankEditMessage}
                                     apiURL={this.props.tenantURL}
                                     embedType={this.props.embedType}
                                     cardId={this.props.cardId}
                                     configButtonLabel='Configure'
                                     assistantButtonLabel='Enable in Digital Assistant'
                                     onConfigure={this._onConfigure} />;
      }
      
    } else if (this.props.displayMode === DisplayMode.Read && (!this.props.tenantURL || !this.props.embedType)) {
      renderPlaceholder = <Placeholder iconName='digital-assistant'
                                     iconText='adenin Digital Assistant'
                                     contentClassName={styles.placeholder}
                                     description={strings.ShowBlankEditMessage}
                                     configButtonLabel='Configure'
                                     assistantButtonLabel='Enable in Digital Assistant'
                                     onConfigure={this._onConfigure} />;
    } else {
      // render the actual card here
      renderCard = <Card {...this.props} />;
    }

    if (this.props.displayMode === DisplayMode.Edit || (this.props.displayMode === DisplayMode.Read && (!this.props.tenantURL || !this.props.embedType))) {
      return (
        
        <div className={ styles.adeninDigitalAssistant }>
          <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} placeholder="Add a title"></WebPartTitle>
          <div className={ styles.container }>
            <ErrorBoundary>
              {renderPlaceholder}
            </ErrorBoundary>
          </div>
        </div>
      );
    } else {
      // render the actual card content
      return (
        
        <div className={ styles.adeninDigitalAssistant }>
          <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} placeholder="Add a title"></WebPartTitle>
          <ErrorBoundary>
            {renderCard}
          </ErrorBoundary>
        </div>
      );
    }
  }

  private _onConfigure = () => {
    this.props.context.propertyPane.open();
  }
}