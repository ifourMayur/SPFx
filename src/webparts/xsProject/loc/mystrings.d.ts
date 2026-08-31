declare interface IXsProjectWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ApiGroupName: string;
  ApiBaseUrlFieldLabel: string;
  ApiBaseUrlFieldDescription: string;
  ApiResourceUriFieldLabel: string;
  ApiResourceUriFieldDescription: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'XsProjectWebPartStrings' {
  const strings: IXsProjectWebPartStrings;
  export = strings;
}
