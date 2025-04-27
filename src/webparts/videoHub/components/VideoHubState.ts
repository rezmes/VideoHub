// VideoHubState.ts - Updated state interface
export interface IVideoHubState {
  filteredVideos: any[];
  searchTerm: string;
  selectedCategory: string;
  selectedDepartment: string;
  durationFilter: string;
  sortBy: string;
  sortDirection: 'asc' | 'desc';
  categoryOptions: {key: string, text: string}[];
  departmentOptions: {key: string, text: string}[];
}