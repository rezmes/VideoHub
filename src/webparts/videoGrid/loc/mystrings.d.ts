// src/webparts/videoGrid/loc/mystrings.d.ts
declare interface IVideoGridWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  LibraryNameFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'VideoGridWebPartStrings' {
  const strings: IVideoGridWebPartStrings;
  export = strings;
}
