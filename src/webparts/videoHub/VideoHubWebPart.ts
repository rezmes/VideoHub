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

export interface IVideoHubWebPartProps {
  libraryName: string;
}

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
            isRTL,
            webUrl: this.context.pageContext.web.absoluteUrl
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

  private _load(): Promise<IVideo[]> {
    const list = this.properties.libraryName || 'KMSVideoHub';
    const currentPageUrl = window.location.href;

    // Query to get video items with author and created date
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('${list}')/items` +
      `?$select=Id,Title,FileRef,FileLeafRef,ThumbnailURL,MediaDuration,Created,Rating` +
      `&$expand=Author` +
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
            : `${this.context.pageContext.web.absoluteUrl}/_layouts/15/next/odspnext/odsp-media/images/itemtypes/96/video.png`;

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
            category: '',  // Will be populated later
            department: '', // Will be populated later
            viewCount: it.Rating || 0,
            uploadDate: this._formatDateAsTimeAgo(new Date(it.Created || new Date())),
            author: it.Author ? it.Author.Title : ''
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
              // Continue even if this fails for one video
            });

          metadataPromises.push(metadataPromise);
        }

        // Wait for all metadata promises to resolve
        return Promise.all(metadataPromises)
          .then(() => {
            return videos;
          })
          .catch(() => {
            return videos;
          });
      });
  }

  // Format date as "time ago" string
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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

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
