import * as React from 'react';
import { IPlaceholderProps } from './IPlaceholderComponent';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Icon } from 'office-ui-fabric-react/lib/components/Icon';
import { registerIcons } from 'office-ui-fabric-react/lib/Styling';
import styles from './Placeholder.module.scss';

export class Placeholder extends React.Component<IPlaceholderProps> {
    private _handleBtnClick = (event?: React.MouseEvent<HTMLButtonElement>): void => {
        this.props.onConfigure();
    }

    private _openAssistantClick = (event?: React.MouseEvent<HTMLButtonElement>): void => {
      // open assistant in a new tab
      window.open('https://app.adenin.com/App/Channels/?name=sharepoint-online-add-in', "_blank");
    }

    private registerAssistantIcon = (): void => {
      registerIcons({
        icons: {
          'digital-assistant': (
            <svg xmlns="http://www.w3.org/2000/svg" id="Layer_1" data-name="Layer 1" viewBox="0 0 190 190">
              <defs>
                <linearGradient id="linear-gradient" x1="136.12" y1="87.85" x2="158.05" y2="105.95" gradientUnits="userSpaceOnUse">
                  <stop offset=".01" stopColor="#817e7c" stopOpacity="0"/>
                  <stop offset=".04" stopColor="#83817f" stopOpacity=".03"/>
                  <stop offset=".07" stopColor="#898887" stopOpacity=".1"/>
                  <stop offset=".1" stopColor="#929495" stopOpacity=".23"/>
                  <stop offset=".14" stopColor="#9fa5a7" stopOpacity=".41"/>
                  <stop offset=".18" stopColor="#afbbbf" stopOpacity=".63"/>
                  <stop offset=".19" stopColor="#b4c2c7" stopOpacity=".7"/>
                  <stop offset=".25" stopColor="#adbbc0" stopOpacity=".77"/>
                  <stop offset=".39" stopColor="#a3b1b6" stopOpacity=".87"/>
                  <stop offset=".54" stopColor="#a0aeb3" stopOpacity=".9"/>
                  <stop offset=".6" stopColor="#a2b0b5" stopOpacity=".86"/>
                  <stop offset=".68" stopColor="#a8b6bb" stopOpacity=".74"/>
                  <stop offset=".78" stopColor="#b3c1c6" stopOpacity=".55"/>
                  <stop offset=".79" stopColor="#b4c2c7" stopOpacity=".52"/>
                  <stop offset="1" stopColor="#fff" stopOpacity="0"/>
                </linearGradient>
                <linearGradient id="linear-gradient-2" x1="36.71" y1="87.34" x2="59.64" y2="105.48" href="#linear-gradient"/>
              </defs>
              <path fill="#f2360a" d="M0 54V0h190V132.67l-44.86-38.04L95 134 45.41 94.63 0 54z"/>
              <path fill="#d50000" d="M190 132.14V190H0V53.5l45.4 40.44 49.59 39.2 50.14-39.2 44.87 38.2z"/>
              <path fill="#FFF" d="M190 123c-9.41-1.85-22-10.46-36.28-28.5-3.55-4.88-7.77-10.4-9.66-12.92C135.9 70.71 115.8 50.64 95 50.64v15c9.68 0 23.81 8.27 40.16 28.88 3.55 4.88 7.77 10.4 9.66 12.92 7.63 10.18 25.79 28.45 45.17 30.69z"/>
              <path d="M159.36 101.25s-3.73-4.33-5.65-6.74c-3.55-4.88-7.77-10.4-9.66-12.92-1-1.39-3.69-4.56-3.69-4.56l-9.43 12.36s2.8 3.32 4.23 5.12c3.55 4.88 7.77 10.4 9.66 12.92 1.38 1.84 5.1 6.16 5.1 6.16z" fill="url(#linear-gradient)"/>
              <path fill="#FFF" d="M190 66c-9.41 1.85-22 10.46-36.28 28.5-3.55 4.88-7.77 10.4-9.66 12.92-8.14 10.85-28.24 30.93-49 30.93v-15c9.68 0 23.81-8.28 40.16-28.89 3.55-4.88 7.77-10.4 9.66-12.92 7.63-10.18 25.79-28.45 45.17-30.69z"/>
              <path fill="#FFF" d="M0 66c9.41 1.85 22 10.46 36.28 28.5 3.55 4.88 7.77 10.4 9.66 12.92 8.14 10.85 28.53 30.92 49.34 30.92v-15c-9.68 0-24.11-8.27-40.45-28.88-3.55-4.88-7.77-10.4-9.66-12.92C37.54 71.4 19.38 53.12 0 50.88z"/>
              <path d="M60 100.66s-3.41-4-5.15-6.18c-3.55-4.88-7.77-10.4-9.66-12.92-1.33-1.78-4.9-5.94-4.9-5.94l-9.5 12.23s3.66 4.25 5.53 6.61c3.55 4.88 7.77 10.4 9.66 12.92 1.19 1.58 4.29 5.25 4.29 5.25z" fill="url(#linear-gradient-2)"/>
              <path fill="#FFF" d="M0 123c9.41-1.85 22-10.46 36.28-28.5 3.55-4.88 7.77-10.4 9.66-12.92C54.1 70.71 74.49 50.64 95.3 50.64v15c-9.68 0-24.11 8.27-40.45 28.88-3.55 4.88-7.77 10.4-9.66 12.92C37.54 117.6 19.38 135.88 0 138.12z"/>
            </svg>
          ),
        },
      });
    }

    public render(): React.ReactElement<IPlaceholderProps> {

      this.registerAssistantIcon();

      return (
        <div className={`${styles.placeholder} ${this.props.contentClassName ? this.props.contentClassName : ''}`}>
          <div className={styles.placeholderContainer}>
            <div className={styles.placeholderHead}>
              <div className={styles.placeholderHeadContainer}>
                {
                  this.props.iconName && <Icon iconName={this.props.iconName} className={styles.placeholderIcon} />
                }
                <span className={styles.placeholderText}>{this.props.iconText}</span>
              </div>
            </div>
            <div className={styles.placeholderDescription}>
              <span className={styles.placeholderDescriptionText}>{this.props.description}</span>
            </div>
            <div className={styles.placeholderBlockDescription}>
              {
                this.props.apiURL && <span className={styles.placeholderDescriptionText}><strong>API URL:</strong> {this.props.apiURL}</span>
              }
              {
                this.props.embedType && <span className={styles.placeholderDescriptionText}><strong>Embed type:</strong> {this.props.embedType}</span>
              }
              {
                this.props.embedType == 'card' && <span className={styles.placeholderDescriptionText}><strong>Card ID:</strong> {this.props.cardId ? this.props.cardId : 'Not set'}</span>
              }
            </div>
            {this.props.children}
            <div className={styles.placeholderDescription}>
              {
                <DefaultButton
                  text={this.props.assistantButtonLabel}
                  ariaLabel={this.props.assistantButtonLabel}
                  onClick={this._openAssistantClick} />
              }
              {
                <PrimaryButton
                  text={this.props.configButtonLabel}
                  ariaLabel={this.props.configButtonLabel}
                  onClick={this._handleBtnClick} />
              }
            </div>
          </div>
        </div>
      );
    }
}