// src/shared/VideoHubStateService.ts
import { IVideo } from './IVideo';

export interface IVideoFilterOptions {
  searchTerm: string;
  selectedCategory: string;
  selectedDepartment: string;
  durationFilter: string;
  sortBy: string;
  sortDirection: 'asc' | 'desc';
}

export class VideoHubStateService {
  private _videos: IVideo[] = [];
  private _filterOptions: IVideoFilterOptions = {
    searchTerm: '',
    selectedCategory: '',
    selectedDepartment: '',
    durationFilter: 'all',
    sortBy: 'title',
    sortDirection: 'asc'
  };
  private _callbacks: Function[] = [];

  // Private constructor to prevent direct instantiation
  private constructor() {
    console.log('VideoHubStateService constructor called');
  }

  public static getInstance(): VideoHubStateService {
    // Use window object to ensure the same instance across web parts
    if (!(window as any).__videoHubStateService) {
      console.log('Creating new VideoHubStateService instance');
      (window as any).__videoHubStateService = new VideoHubStateService();
    } else {
      console.log('Reusing existing VideoHubStateService instance');
    }
    return (window as any).__videoHubStateService;
  }

  public getVideos(): IVideo[] {
    return this._videos;
  }

  public setVideos(videos: IVideo[]): void {
    console.log('Setting videos:', videos.length);
    this._videos = videos;
    this._notifyCallbacks();
  }

  public getFilterOptions(): IVideoFilterOptions {
    return this._filterOptions;
  }

  public updateFilterOptions(newOptions: Partial<IVideoFilterOptions>): void {
    console.log('Updating filter options:', newOptions);

    // Create a new object by merging the current options with the new options
    this._filterOptions = {
      searchTerm: newOptions.searchTerm !== undefined ? newOptions.searchTerm : this._filterOptions.searchTerm,
      selectedCategory: newOptions.selectedCategory !== undefined ? newOptions.selectedCategory : this._filterOptions.selectedCategory,
      selectedDepartment: newOptions.selectedDepartment !== undefined ? newOptions.selectedDepartment : this._filterOptions.selectedDepartment,
      durationFilter: newOptions.durationFilter !== undefined ? newOptions.durationFilter : this._filterOptions.durationFilter,
      sortBy: newOptions.sortBy !== undefined ? newOptions.sortBy : this._filterOptions.sortBy,
      sortDirection: newOptions.sortDirection !== undefined ? newOptions.sortDirection : this._filterOptions.sortDirection
    };

    console.log('Updated filter options:', this._filterOptions);
    this._notifyCallbacks();
  }

  public subscribe(callback: Function): void {
    // Store callback with a unique identifier
    const callbackId = 'callback_' + new Date().getTime() + '_' + Math.random();
    console.log('Adding subscriber with ID:', callbackId);

    this._callbacks.push(callback);
    console.log('Subscriber added, total subscribers:', this._callbacks.length);

    // Log all subscribers for debugging
    console.log('Current subscribers:', this._callbacks);
  }

  public unsubscribe(callback: Function): void {
    const index = this._callbacks.indexOf(callback);
    if (index !== -1) {
      this._callbacks.splice(index, 1);
      console.log('Subscriber removed, total subscribers:', this._callbacks.length);
    }
  }

  private _notifyCallbacks(): void {
    console.log('Notifying', this._callbacks.length, 'subscribers');
    for (let i = 0; i < this._callbacks.length; i++) {
      try {
        console.log('Calling subscriber', i);
        this._callbacks[i]();
      } catch (error) {
        console.error('Error in subscriber callback:', error);
      }
    }
  }

  // Add to VideoHubStateService.ts
// Add to VideoHubStateService.ts
public getFilteredVideos(): IVideo[] {
  const videos = this._videos;
  const filterOptions = this._filterOptions;
  let filtered = videos.slice();

  // Filter by search term
  if (filterOptions.searchTerm) {
    const term = filterOptions.searchTerm.toLowerCase();
    filtered = filtered.filter(video => 
      video.title.toLowerCase().indexOf(term) !== -1
    );
  }

  // Filter by category
  if (filterOptions.selectedCategory) {
    filtered = filtered.filter(video => 
      video.category === filterOptions.selectedCategory
    );
  }

  // Filter by department
  if (filterOptions.selectedDepartment) {
    filtered = filtered.filter(video => 
      video.department === filterOptions.selectedDepartment
    );
  }

  // Filter by duration
  if (filterOptions.durationFilter !== "all") {
    filtered = filtered.filter(video => {
      const durationInSeconds = this._parseDuration(video.duration);
      
      if (filterOptions.durationFilter === "short") {
        return durationInSeconds < 300; // Less than 5 minutes
      } else if (filterOptions.durationFilter === "medium") {
        return durationInSeconds >= 300 && durationInSeconds < 900; // 5-15 minutes
      } else if (filterOptions.durationFilter === "long") {
        return durationInSeconds >= 900; // 15+ minutes
      }
      
      return true;
    });
  }

  // Sort videos
  filtered.sort((a, b) => {
    let comparison = 0;

    if (filterOptions.sortBy === "title") {
      comparison = a.title.localeCompare(b.title);
    } else if (filterOptions.sortBy === "duration") {
      comparison = this._parseDuration(a.duration) - this._parseDuration(b.duration);
    } else if (filterOptions.sortBy === "category") {
      comparison = (a.category || "").localeCompare(b.category || "");
    } else if (filterOptions.sortBy === "department") {
      comparison = (a.department || "").localeCompare(b.department || "");
    }

    return filterOptions.sortDirection === "asc" ? comparison : -comparison;
  });

  return filtered;
}

private _parseDuration(duration: string): number {
  if (!duration) return 0;

  const parts = duration.split(":").map(part => parseInt(part, 10));

  if (parts.length === 2) {
    return parts[0] * 60 + parts[1];
  } else if (parts.length === 3) {
    return parts[0] * 3600 + parts[1] * 60 + parts[2];
  }

  return 0;
}



}
