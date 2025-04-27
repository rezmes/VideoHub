// src/shared/VideoLoaderService.ts
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IVideo } from './IVideo';

export class VideoLoaderService {
  private static instance: VideoLoaderService;

  private constructor() {
    console.log('VideoLoaderService constructor called');
  }

  public static getInstance(): VideoLoaderService {
    if (!VideoLoaderService.instance) {
      VideoLoaderService.instance = new VideoLoaderService();
    }
    return VideoLoaderService.instance;
  }

// In VideoLoaderService.ts
public loadVideos(
  webUrl: string,
  listName: string,
  spHttpClient: SPHttpClient,
  currentPageUrl: string
): Promise<IVideo[]> {
  // Very simple query - just get the basic fields
  const url =
  `${webUrl}` +
  `/_api/web/lists/getbytitle('${listName}')/items` +
  `?$select=Id,Title,FileRef,FileLeafRef,ThumbnailURL,MediaDuration,Created,Rating` +
  `&$orderby=Modified desc&$top=50`;


  console.log('Fetching videos with URL:', url);

  return spHttpClient
    .get(url, SPHttpClient.configurations.v1)
    .then((r: SPHttpClientResponse) => {
      if (!r.ok) {
        console.error('Error response:', r.status, r.statusText);
        throw new Error(`HTTP ${r.status}`);
      }
      return r.json();
    })
    // Rest of the method...
}

// In VideoLoaderService.ts
public loadVideosSimple(
  webUrl: string,
  listName: string,
  spHttpClient: SPHttpClient,
  currentPageUrl: string
): Promise<IVideo[]> {
  // Very simple query - just get the basic fields
  const url =
    `${webUrl}` +
    `/_api/web/lists/getbytitle('${listName}')/items` +
    `?$select=Id,Title,FileRef,FileLeafRef` +
    `&$orderby=Modified desc&$top=10`;

  console.log('Fetching videos with URL:', url);

  return spHttpClient
    .get(url, SPHttpClient.configurations.v1)
    .then((r: SPHttpClientResponse) => {
      if (!r.ok) {
        console.error('Error response:', r.status, r.statusText);
        throw new Error(`HTTP ${r.status}`);
      }
      return r.json();
    })
    .then(j => {
      console.log('Got response, processing items...');
      // Filter MP4 files client-side
      const items = j.value || [];
      const mp4Items = [];

      // Filter using traditional loop
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        if (item.FileLeafRef && item.FileLeafRef.toLowerCase().indexOf('.mp4') !== -1) {
          mp4Items.push(item);
        }
      }

      console.log('Found MP4 items:', mp4Items.length);

      // Map to IVideo objects with minimal fields
      const videos: IVideo[] = [];
      for (let i = 0; i < mp4Items.length; i++) {
        const it = mp4Items[i];

        // Create the modern player URL
        const fileRef = it.FileRef || '';
        const encodedFileRef = encodeURIComponent(fileRef);
        const parentFolder = fileRef.substring(0, fileRef.lastIndexOf('/'));
        const encodedParentFolder = encodeURIComponent(parentFolder);

        const modernPlayerUrl =
          `${parentFolder}/Forms/AllItems.aspx?Source=${encodeURIComponent(currentPageUrl)}` +
          `&id=${encodedFileRef}&parent=${encodedParentFolder}`;

        videos.push({
          id: it.Id,
          title: it.Title || (it.FileLeafRef ? it.FileLeafRef.replace('.mp4', '') : 'Untitled Video'),
          url: modernPlayerUrl,
          previewUrl: `${webUrl}/_layouts/15/next/odspnext/odsp-media/images/itemtypes/96/video.png`,
          duration: '',
          category: '',
          department: '',
          viewCount: 0,
          uploadDate: 'Unknown',
          author: ''
        });
      }

      console.log('Processed videos:', videos.length);
      return videos;
    });
}
// In VideoLoaderService.ts
public loadVideosWithThumbnails(
  webUrl: string,
  listName: string,
  spHttpClient: SPHttpClient,
  currentPageUrl: string
): Promise<IVideo[]> {
  // Add ThumbnailURL to the query
  const url =
    `${webUrl}` +
    `/_api/web/lists/getbytitle('${listName}')/items` +
    `?$select=Id,Title,FileRef,FileLeafRef,ThumbnailURL,MediaDuration` +
    `&$orderby=Modified desc&$top=10`;

  console.log('Fetching videos with thumbnails URL:', url);

  return spHttpClient
    .get(url, SPHttpClient.configurations.v1)
    .then((r: SPHttpClientResponse) => {
      if (!r.ok) {
        console.error('Error response:', r.status, r.statusText);
        throw new Error(`HTTP ${r.status}`);
      }
      return r.json();
    })
    .then(j => {
      console.log('Got response with thumbnails, processing items...');
      // Filter MP4 files client-side
      const items = j.value || [];
      const mp4Items = [];

      // Filter using traditional loop
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        if (item.FileLeafRef && item.FileLeafRef.toLowerCase().indexOf('.mp4') !== -1) {
          mp4Items.push(item);
        }
      }

      console.log('Found MP4 items:', mp4Items.length);

      // Map to IVideo objects with thumbnails
      const videos: IVideo[] = [];
      for (let i = 0; i < mp4Items.length; i++) {
        const it = mp4Items[i];

        // Get thumbnail URL from ThumbnailURL field
        const thumb = it.ThumbnailURL && it.ThumbnailURL.Url
          ? it.ThumbnailURL.Url
          : `${webUrl}/_layouts/15/next/odspnext/odsp-media/images/itemtypes/96/video.png`;

        // Create the modern player URL
        const fileRef = it.FileRef || '';
        const encodedFileRef = encodeURIComponent(fileRef);
        const parentFolder = fileRef.substring(0, fileRef.lastIndexOf('/'));
        const encodedParentFolder = encodeURIComponent(parentFolder);

        const modernPlayerUrl =
          `${parentFolder}/Forms/AllItems.aspx?Source=${encodeURIComponent(currentPageUrl)}` +
          `&id=${encodedFileRef}&parent=${encodedParentFolder}`;

        videos.push({
          id: it.Id,
          title: it.Title || (it.FileLeafRef ? it.FileLeafRef.replace('.mp4', '') : 'Untitled Video'),
          url: modernPlayerUrl,
          previewUrl: thumb,
          duration: it.MediaDuration || '',
          category: '',
          department: '',
          viewCount: 0,
          uploadDate: 'Unknown',
          author: ''
        });
      }

      console.log('Processed videos with thumbnails:', videos.length);
      return videos;
    });
}
// In VideoLoaderService.ts
// In VideoLoaderService.ts
public loadVideosWithCreatedAndAuthor(
  webUrl: string,
  listName: string,
  spHttpClient: SPHttpClient,
  currentPageUrl: string
): Promise<IVideo[]> {
  // Add Created and Author fields, but not Rating
  const url =
    `${webUrl}` +
    `/_api/web/lists/getbytitle('${listName}')/items` +
    `?$select=Id,Title,FileRef,FileLeafRef,ThumbnailURL,MediaDuration,Created,Author/Title` +
    `&$expand=Author` +
    `&$orderby=Modified desc&$top=10`;

  console.log('Fetching videos with Created and Author URL:', url);

  return spHttpClient
    .get(url, SPHttpClient.configurations.v1)
    .then((r: SPHttpClientResponse) => {
      if (!r.ok) {
        console.error('Error response:', r.status, r.statusText);
        throw new Error(`HTTP ${r.status}`);
      }
      return r.json();
    })
    .then(j => {
      console.log('Got response with Created and Author, processing items...');
      // Filter MP4 files client-side
      const items = j.value || [];
      const mp4Items = [];

      // Filter using traditional loop
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        if (item.FileLeafRef && item.FileLeafRef.toLowerCase().indexOf('.mp4') !== -1) {
          mp4Items.push(item);
        }
      }

      console.log('Found MP4 items:', mp4Items.length);

      // Map to IVideo objects with thumbnails, Created date, and Author
      const videos: IVideo[] = [];
      for (let i = 0; i < mp4Items.length; i++) {
        const it = mp4Items[i];

        // Get thumbnail URL from ThumbnailURL field
        const thumb = it.ThumbnailURL && it.ThumbnailURL.Url
          ? it.ThumbnailURL.Url
          : `${webUrl}/_layouts/15/next/odspnext/odsp-media/images/itemtypes/96/video.png`;

        // Create the modern player URL
        const fileRef = it.FileRef || '';
        const encodedFileRef = encodeURIComponent(fileRef);
        const parentFolder = fileRef.substring(0, fileRef.lastIndexOf('/'));
        const encodedParentFolder = encodeURIComponent(parentFolder);

        const modernPlayerUrl =
          `${parentFolder}/Forms/AllItems.aspx?Source=${encodeURIComponent(currentPageUrl)}` +
          `&id=${encodedFileRef}&parent=${encodedParentFolder}`;

        // Format the Created date as "time ago"
        const uploadDate = this._formatDateAsTimeAgo(new Date(it.Created || new Date()));

        videos.push({
          id: it.Id,
          title: it.Title || (it.FileLeafRef ? it.FileLeafRef.replace('.mp4', '') : 'Untitled Video'),
          url: modernPlayerUrl,
          previewUrl: thumb,
          duration: it.MediaDuration || '',
          category: '',  // Will be populated later
          department: '', // Will be populated later
          viewCount: 0,  // Use a default value instead of Rating
          uploadDate: uploadDate,
          author: it.Author ? it.Author.Title : ''
        });
      }

      // Now fetch metadata for each video using FieldValuesAsText
      const metadataPromises = [];

      for (let i = 0; i < videos.length; i++) {
        const video = videos[i];

        // Create a promise to get the formatted field values
        const metadataPromise = this._getFieldValuesAsText(webUrl, listName, video.id, spHttpClient)
          .then(fieldValues => {
            // Extract category and department from the field values
            if (fieldValues.CategoryMangedMetaData) {
              video.category = fieldValues.CategoryMangedMetaData;
            }

            if (fieldValues.DepartmentM) {
              video.department = fieldValues.DepartmentM;
            }
          })
          .catch(err => {
            console.log(`Error getting metadata for video ${video.id}:`, err);
            // Continue even if this fails for one video
          });

        metadataPromises.push(metadataPromise);
      }

      // Wait for all metadata promises to resolve
      return Promise.all(metadataPromises)
        .then(() => {
          console.log('Processed videos with metadata:', videos.length);
          return videos;
        })
        .catch(err => {
          console.log('Error fetching metadata:', err);
          return videos;
        });
    });
}

// Helper method to get formatted field values
private _getFieldValuesAsText(
  webUrl: string,
  listTitle: string,
  itemId: number,
  spHttpClient: SPHttpClient
): Promise<any> {
  const url =
    `${webUrl}` +
    `/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})/FieldValuesAsText`;

  return spHttpClient
    .get(url, SPHttpClient.configurations.v1)
    .then(r => {
      if (!r.ok) {
        throw new Error(`HTTP ${r.status}`);
      }
      return r.json();
    })
    .then(result => {
      return result || {};
    });
}


// Add this helper method to format dates as "time ago"
private _formatDateAsTimeAgo(date: Date): string {
  const now = new Date();
  const diffInMs = now.getTime() - date.getTime();
  const diffInDays = Math.floor(diffInMs / (1000 * 60 * 60 * 24));

  if (diffInDays < 1) {
    return 'Today';
  } else if (diffInDays === 1) {
    return 'Yesterday';
  } else if (diffInDays < 7) {
    return `${diffInDays} days ago`;
  } else if (diffInDays < 30) {
    const weeks = Math.floor(diffInDays / 7);
    return `${weeks} ${weeks === 1 ? 'week' : 'weeks'} ago`;
  } else if (diffInDays < 365) {
    const months = Math.floor(diffInDays / 30);
    return `${months} ${months === 1 ? 'month' : 'months'} ago`;
  } else {
    const years = Math.floor(diffInDays / 365);
    return `${years} ${years === 1 ? 'year' : 'years'} ago`;
  }
}



  // Get field values as text for metadata

}
