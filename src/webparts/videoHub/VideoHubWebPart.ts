
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import VideoHub from './components/VideoHub';
import { IVideo } from '../../shared/IVideo';

export interface IVideoHubWebPartProps { libraryName: string; }

export default class VideoHubWebPart
  extends BaseClientSideWebPart<IVideoHubWebPartProps> {

    public render(): void {
      // Show loading indicator
      ReactDom.render(
        React.createElement('div', {}, 'Loading videos...'),
        this.domElement
      );

      // Detect if the page is in RTL mode
      const isRTL = document.documentElement.getAttribute('dir') === 'rtl';

      this._load()
        .then(videos =>
          ReactDom.render(
            React.createElement(VideoHub, {
              videos,
              isRTL
            }),
            this.domElement
          )
        )
        .catch(err => {
          console.error('Error loading videos:', err);
          ReactDom.render(
            React.createElement('p', {}, `⚠️ ${err.message}`),
            this.domElement
          );
        });
    }


  /* ---------- REST + thumbnail picker ---------- */
  // VideoHubWebPart.ts - Update the _load method to include metadata fields
// VideoHubWebPart.ts - Update the _load method
// private _load(): Promise<IVideo[]> {
//   const list = this.properties.libraryName || 'KMSVideoHub';

//   // Get the current page URL for the return link
//   const currentPageUrl = window.location.href;

//   // Simplified query without managed metadata fields
//   const url =
//     `${this.context.pageContext.web.absoluteUrl}` +
//     `/_api/web/lists/getbytitle('${list}')/items` +
//     `?$select=Id,Title,FileRef,FileLeafRef,ThumbnailURL,MediaDuration` +
//     `&$orderby=Modified desc&$top=50`;

//   return this.context.spHttpClient
//     .get(url, SPHttpClient.configurations.v1)
//     .then((r: SPHttpClientResponse) => {
//       if (!r.ok) { throw new Error(`HTTP ${r.status}`); }
//       return r.json();
//     })
//     .then(j => {
//       // Log the first item to see its structure
//       if (j.value && j.value.length > 0) {
//         console.log('First item structure:', j.value[0]);
//       }

//       // Filter MP4 files client-side
//       const items = j.value || [];
//       const mp4Items = items.filter(function(item) {
//         return item.FileLeafRef && item.FileLeafRef.toLowerCase().endsWith('.mp4');
//       });

//       return mp4Items.map(function(it) {
//         // Get thumbnail URL from ThumbnailURL field
//         const thumb = it.ThumbnailURL && it.ThumbnailURL.Url
//           ? it.ThumbnailURL.Url
//           : `${this.context.pageContext.web.absoluteUrl}/_layouts/15/images/videoicon.png`;

//         // Create the modern player URL based on the example
//         const fileRef = it.FileRef || '';

//         const encodedFileRef = encodeURIComponent(fileRef);
//         const parentFolder = fileRef.substring(0, fileRef.lastIndexOf('/'));
//         const encodedParentFolder = encodeURIComponent(parentFolder);
//         const encodedSourceUrl = encodeURIComponent(currentPageUrl);

//         const modernPlayerUrl =
//           `${parentFolder}/Forms/AllItems.aspx?Source=${encodeURIComponent(currentPageUrl)}` +
//           `&id=${encodedFileRef}&parent=${encodedParentFolder}`;

//         // For now, use empty strings for category and department
//         return {
//           id: it.Id,
//           title: it.Title || (it.FileLeafRef ? it.FileLeafRef.replace('.mp4', '') : 'Untitled Video'),
//           url: modernPlayerUrl,
//           previewUrl: thumb,
//           duration: it.MediaDuration || '',
//           category: '',  // Temporarily empty
//           department: '' // Temporarily empty
//         } as IVideo;
//       }, this);
//     });
// }
private _getTermLabel(termId: string): Promise<string> {
  if (!termId || typeof termId !== 'string') {
    return Promise.resolve('');
  }

  // If the termId contains a pipe character, extract the label directly
  if (termId.indexOf('|') > -1) {
    const parts = termId.split('|');
    if (parts.length > 1 && parts[1]) {
      return Promise.resolve(parts[1]);
    }
  }

  // If we couldn't extract the label directly, try to get it from the hidden list
  // This is a fallback and might not be needed if the format is always "GUID|Label"
  const guid = termId.split('|')[0];

  return this.context.spHttpClient
    .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('TaxonomyHiddenList')/items?$filter=IdForTerm eq '${guid}'`,
      SPHttpClient.configurations.v1)
    .then((r: SPHttpClientResponse) => {
      if (!r.ok) {
        // If this fails, just return empty string rather than failing the whole operation
        console.log('Error fetching term label:', r.status);
        return { value: [] };
      }
      return r.json();
    })
    .then(j => {
      if (j.value && j.value.length > 0) {
        return j.value[0].Term || '';
      }
      return '';
    })
    .catch(err => {
      console.log('Error in _getTermLabel:', err);
      return '';
    });
}


// Replace the entire metadata fetching part with this simpler approach
private _load(): Promise<IVideo[]> {
  const list = this.properties.libraryName || 'KMSVideoHub';

  // Get the current page URL for the return link
  const currentPageUrl = window.location.href;

  // Very simple query - just get the basic fields without trying to expand metadata
  const url =
    `${this.context.pageContext.web.absoluteUrl}` +
    `/_api/web/lists/getbytitle('${list}')/items` +
    `?$select=Id,Title,FileRef,FileLeafRef,ThumbnailURL,MediaDuration` +
    `&$orderby=Modified desc&$top=50`;

  return this.context.spHttpClient
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
          : `${this.context.pageContext.web.absoluteUrl}/_layouts/15/images/videoicon.png`;

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
        const metadataPromise = this._getFieldValuesAsText(list, video.id)
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

// Helper method to get formatted field values
private _getFieldValuesAsText(listTitle: string, itemId: number): Promise<any> {
  const url =
    `${this.context.pageContext.web.absoluteUrl}` +
    `/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})/FieldValuesAsText`;

  return this.context.spHttpClient
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


// Helper method to get list fields
private _getListFields(listTitle: string): Promise<any[]> {
  const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/fields?$filter=Hidden eq false`;

  return this.context.spHttpClient
    .get(url, SPHttpClient.configurations.v1)
    .then(r => r.json())
    .then(result => result.value || [])
    .catch(err => {
      console.log('Error fetching list fields:', err);
      return [];
    });
}

// Helper method to extract term label from taxonomy field value
private _extractTermLabel(value: any): string {
  if (!value) {
    return '';
  }

  // Check for the text companion field format (ends with _0)
  for (const key in value) {
    if (key.indexOf('_0', key.length - 2) !== -1 && typeof value[key] === 'string') {
      return value[key];
    }
  }

  // If it's a string with a pipe character, extract the label
  if (typeof value === 'string' && value.indexOf('|') > -1) {
    return value.split('|')[1] || '';
  }

  // If it's an object with a Label property
  if (typeof value === 'object' && value.Label) {
    return value.Label;
  }

  // If it's an object with a Term property
  if (typeof value === 'object' && value.Term) {
    return value.Term;
  }

  // If we can't extract it, return the string representation
  return String(value);
}

  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'Settings' },
        groups: [{
          groupFields: [
            PropertyPaneTextField('libraryName', {
              label: 'Library title',
              value: 'KMSVideoHub'
            })
          ]
        }]
      }]
    };
  }
}
