declare interface IAssistantCardsWebPartStrings {
  leaveBlankForDefault: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  componentConfigName: string;
  componentClientIDName: string;
  tenantURLLabel: string;
  componentCDNFieldLabel: string;
  componentClientIDFieldLabel: string;
  embedTypeDropdownLabel: string;
  contextLoaderLabel: string;
  customCSSLabel: string;
  customCSSDescription: string;
  embedTypeName: string;
  defaultCDN: string;
  defaultClientID: string;
  defaultContextLoader: string;
}

declare module 'AssistantCardsWebPartStrings' {
  const strings: IAssistantCardsWebPartStrings;
  export = strings;
}
