// src/webparts/videoFilter/loc/mystrings.d.ts
declare interface IVideoFilterWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  LibraryNameFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'VideoFilterWebPartStrings' {
  const strings: IVideoFilterWebPartStrings;
  export = strings;
}
