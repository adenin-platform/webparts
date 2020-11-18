import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IAdeninDigitalAssistantProps {
  tenantURL: string;
  componentCDN: string;
  SSOProviderID: string;
  componentClientID: string;
  embedType: string;
  cardId: string;
  cardStyle: string;
  customCSSClasses: string;
  title: string;
  displayMode: DisplayMode;
  context: WebPartContext;
  tenantId: string;
  updateProperty: (value: string) => void;
}