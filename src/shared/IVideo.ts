// IVideo.ts - Update to include metadata fields
// src/shared/IVideo.ts
export interface IVideo {
  id: number;
  title: string;
  url: string;
  previewUrl: string;
  duration: string;
  category: string;
  department: string;
  viewCount?: number;
  uploadDate?: string;
  author?: string;  // Add this line
}
