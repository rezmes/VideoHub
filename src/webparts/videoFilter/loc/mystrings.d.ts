// // src/webparts/videoFilter/loc/mystrings.d.ts
// declare interface IVideoFilterWebPartStrings {
//   PropertyPaneDescription: string;
//   BasicGroupName: string;
//   LibraryNameFieldLabel: string;
//   AppLocalEnvironmentSharePoint: string;
//   AppLocalEnvironmentTeams: string;
//   AppSharePointEnvironment: string;
//   AppTeamsTabEnvironment: string;
// }

// declare module 'VideoFilterWebPartStrings' {
//   const strings: IVideoFilterWebPartStrings;
//   export = strings;
// }

// src/webparts/videoFilter/loc/mystrings.d.ts
declare interface IVideoFilterWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  
  // UI strings
  FilterTitle: string;
  ResetFilters: string;
  SearchPlaceholder: string;
  CategoryLabel: string;
  DepartmentLabel: string;
  DurationLabel: string;
  AllDurations: string;
  ShortVideos: string;
  MediumVideos: string;
  LongVideos: string;
  SortByLabel: string;
  SortDirectionAsc: string;
  SortDirectionDesc: string;
}

declare module 'VideoFilterWebPartStrings' {
  const strings: IVideoFilterWebPartStrings;
  export = strings;
}
