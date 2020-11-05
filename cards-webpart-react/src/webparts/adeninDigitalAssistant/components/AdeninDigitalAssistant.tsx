import * as React from 'react';
import styles from './AdeninDigitalAssistant.module.scss';
import { IAdeninDigitalAssistantProps } from './IAdeninDigitalAssistantProps';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Text } from 'office-ui-fabric-react/lib/Text';
import { Placeholder } from "./Placeholder/Placeholder";
import { DisplayMode } from '@microsoft/sp-core-library';
import * as strings from 'AdeninDigitalAssistantWebPartStrings';

export default class AdeninDigitalAssistant extends React.Component<IAdeninDigitalAssistantProps, {}> {
  public render(): React.ReactElement<IAdeninDigitalAssistantProps> {
    let renderContent: JSX.Element = null;

    if (this.props.displayMode === DisplayMode.Edit && (!this.props.tenantURL || !this.props.embedType)) {
      renderContent = <Placeholder iconName='digital-assistant'
                                   iconText='adenin Digital Assistant'
                                   contentClassName={styles.placeholder}
                                   description={strings.ShowBlankEditMessage}
                                   configButtonLabel='Configure'
                                   assistantButtonLabel='Open Digital Assistant'
                                   onConfigure={this._onConfigure} />;
    } else if (this.props.displayMode === DisplayMode.Edit) {
      renderContent = <Placeholder iconName='digital-assistant'
                        iconText='adenin Digital Assistant'
                        contentClassName={styles.placeholder}
                        description={strings.ShowBlankEditMessage}
                        apiURL={this.props.tenantURL}
                        embedType={this.props.embedType}
                        cardId={this.props.cardId}
                        configButtonLabel='Configure'
                        assistantButtonLabel='Open Digital Assistant'
                        onConfigure={this._onConfigure} />;
    } else {
      // render the actual card here
      renderContent = <Text>Hi, I'll be a card!</Text>;
    }

    if (this.props.displayMode === DisplayMode.Edit) {
      return (
        <div className={ styles.adeninDigitalAssistant }>
          <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} placeholder="Add a title"></WebPartTitle>
          <div className={ styles.container }>
            {renderContent}
          </div>
        </div>
      );
    } else {
      // render the actual card content
      return (
        <div className={ styles.adeninDigitalAssistant }>
          <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} placeholder="Add a title"></WebPartTitle>
          {renderContent}
        </div>
      );
    }
  }

  private _onConfigure = () => {
    this.props.context.propertyPane.open();
  }
}