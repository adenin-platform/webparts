import { DisplayMode } from '@microsoft/sp-core-library';

export interface IAdeninDigitalAssistantProps {
  tenantURL: string;
  componentCDN: string;
  SSOProviderID: string;
  componentClientID: string;
  contextLoaderSrc: string;
  embedType: string;
  cardId: string;
  cardStyle: string;
  customCSSClasses: string;
  title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
}