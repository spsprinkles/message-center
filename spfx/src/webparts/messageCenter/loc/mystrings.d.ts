declare interface IMessageCenterWebPartStrings {
  ListNameFieldDescription: string;
  ListNameFieldLabel: string;
  MoreInfoFieldDescription: string;
  MoreInfoFieldLabel: string;
  MoreInfoTooltipFieldDescription: string;
  MoreInfoTooltipFieldLabel: string;
  TileColumnSizeFieldLabel: string;
  TilePageSizeFieldLabel: string;
  TimeFormatFieldDescription: string;
  TimeFormatFieldLabel: string;
  TimeZoneFieldLabel: string;
  TitleFieldDescription: string;
  TitleFieldLabel: string;
  //SortFieldDescription: string;
  SortFieldLabel: string;
  WebUrlFieldDescription: string;
  WebUrlFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
}

declare module 'MessageCenterWebPartStrings' {
  const strings: IMessageCenterWebPartStrings;
  export = strings;
}
