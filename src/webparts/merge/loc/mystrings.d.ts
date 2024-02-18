declare interface IMergeWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  VisibilityGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  WhoCanSee: string;
  WhoCanEdit: string;
  Everyone: string;
  Admins:string;

  SiteName: string;
  ListFor: string;
  ColumnFor:string;
  ShowListData: string;
  // Panel 
  MergeLists:string;
  ListOptionsFieldLabel: string;
  VisibilityOptionUsers: string;
  VisibilityOptionAdministrators: string;
  VisibilityOptionHide: string;
  VisibilitySaving: string;
  VisibilitySaved: string;

  SelectSite: string;
  SelectList: string;
  SelectColumn: string;

  // Merge
  MergeSettings: string;
  CloseButtonText: string;
  ButtonText: string;
  ButtonIconName: string;
  ButtonAlignment: string;
  ButtonAlignmentLeft: string;
  ButtonAlignmentCenter: string;
  ButtonAlignmentRight: string;
  ButtonSize: string;
  ButtonShowAs: string;
  ButtonShowAsButton: string;
  ButtonShowAsLink: string;

  //LoadGrid
  FilterLabel: string;
  FilterTextboxPlaceholder: string;
  SelectColumn: string;
  AccessDenied: string;
}

declare module 'MergeWebPartStrings' {
  const strings: IMergeWebPartStrings;
  export = strings;
}
