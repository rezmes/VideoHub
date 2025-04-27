// IVideo.ts - Keep this file as is

// VideoHubState.ts - New file
export interface IVideoHubState {
  filteredVideos: any[];
  searchTerm: string;
  sortBy: string;
  sortDirection: 'asc' | 'desc';
}

