declare interface IMetricsWebPartStrings {
  PropertyPaneDescription: string;
  ConfigGroupName: string;
  ConfigSourceLabel: string;
  ConfigSourceInline: string;
  ConfigSourceUrl: string;
  OpenEditorLabel: string;
  ProblemsLabel: string;
  ConfigJsonLabel: string;
  ConfigUrlLabel: string;
  ConfigUrlDescription: string;
  ReloadLabel: string;
  LoadFailed: string;
  InvalidJson: string;
  WebPartProperties: string;
}

declare module 'MetricsWebPartStrings' {
  const strings: IMetricsWebPartStrings;
  export = strings;
}