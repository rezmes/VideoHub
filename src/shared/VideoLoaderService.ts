// src/shared/VideoLoaderService.ts
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IVideo } from './IVideo';

export class VideoLoaderService {
  private static instance: VideoLoaderService;

  private constructor() {}

  public static getInstance(): VideoLoaderService {
    if (!VideoLoaderService.instance) {
      VideoLoaderService.instance = new VideoLoaderService();
    }
    return VideoLoaderService.instance;
  }

  public loadVideos(
    webUrl: string,
    listName: string,
    spHttpClient: SPHttpClient,
    currentPageUrl: string
  ): Promise<IVideo[]> {
    // Very simple query - just get the basic fields without trying to expand metadata
    const url =
      `${webUrl}` +
      `/_api/web/lists/getbytitle('${listName}')/items` +
      `?$select=Id,Title,FileRef,FileLeafRef,ThumbnailURL,MediaDuration` +
      `&$orderby=Modified desc&$top=50`;

    return spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => {
        if (!r.ok) { throw new Error(`HTTP ${r.status}`); }
        return r.json();
      })
      .then(j => {
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

        // Map to IVideo objects without metadata for now
        const videos: IVideo[] = [];
        for (let i = 0; i < mp4Items.length; i++) {
          const it = mp4Items[i];

          // Get thumbnail URL from ThumbnailURL field
          const thumb = it.ThumbnailURL && it.ThumbnailURL.Url
            ? it.ThumbnailURL.Url
            : `${webUrl}/_layouts/15/images/videoicon.png`;

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
            category: '', // Empty for now
            department: '' // Empty for now
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
            return videos;
          })
          .catch(err => {
            console.log('Error fetching metadata:', err);
            return videos;
          });
      });
  }

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
}
