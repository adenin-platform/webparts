declare interface IAssistantCardsWebPartStrings {
  leaveBlankForDefault: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  componentConfigName: string;
  componentClientIDName: string;
  APIEndpointFieldLabel: string;
  componentCDNFieldLabel: string;
  componentClientIDFieldLabel: string;
  embedTypeDropdownLabel: string;
  embedTypeName: string;
  defaultCDN: string;
  defaultClientID: string;
}

declare module 'AssistantCardsWebPartStrings' {
  const strings: IAssistantCardsWebPartStrings;
  export = strings;
}
