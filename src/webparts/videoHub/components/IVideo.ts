// IVideo.ts - Update to include metadata fields
export interface IVideo {
  id: number;
  title: string;
  url: string;
  previewUrl: string;
  duration: string;
  category: string;  // For CategoryMangedMetaData
  department: string; // For DepartmentM
}