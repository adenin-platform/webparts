import * as React from 'react';
import styles from './AdeninDigitalAssistant.module.scss';
import { IAdeninDigitalAssistantProps } from './IAdeninDigitalAssistantProps';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { escape } from '@microsoft/sp-lodash-subset';
import { DisplayMode } from '@microsoft/sp-core-library';
import * as strings from 'AdeninDigitalAssistantWebPartStrings';

export default class AdeninDigitalAssistant extends React.Component<IAdeninDigitalAssistantProps, {}> {
  public render(): React.ReactElement<IAdeninDigitalAssistantProps> {
    let renderContent: JSX.Element = null;

    if (this.props.displayMode === DisplayMode.Edit && !this.props.tenantURL) {
      renderContent = <Card cardType={cardType.info}>{strings.ShowBlankEditMessage}</Card>;
    } else if (this.props.displayMode === DisplayMode.Edit) {
      renderContent = <Card cardType={cardType.info} tenantURL={this.props.tenantURL} embedType={this.props.embedType}/>;
    } else {
      // render the actual card here
    }

    if (this.props.displayMode === DisplayMode.Edit) {
      return (
        <div className={ styles.adeninDigitalAssistant }>
          <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} placeholder="Add a title"></WebPartTitle>
          <div className={ styles.container }>
            <div className={ styles.row }>
              <div className={ styles.column }>
                <span className={ styles.title }>adenin Digital Assistant</span>
                {renderContent};
              </div>
            </div>
          </div>
        </div>
      );
    } else {
      // render the actual card content
      return ( 
        <div className={ styles.adeninDigitalAssistant }>
          <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} placeholder="Add a title"></WebPartTitle>);
          {renderContent};
        </div>
      );
    }
  }
}
