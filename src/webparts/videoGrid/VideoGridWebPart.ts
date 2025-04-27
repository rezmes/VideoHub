// src/webparts/videoGrid/VideoGridWebPart.ts
// At the top of your VideoGridWebPart.ts file, add this import:
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import VideoGrid from './components/VideoGrid';
import { VideoHubStateService } from '../../shared/VideoHubStateService';
import { VideoLoaderService } from '../../shared/VideoLoaderService';

export interface IVideoGridWebPartProps {
  libraryName: string;

}

export default class VideoGridWebPart extends BaseClientSideWebPart<IVideoGridWebPartProps> {
  private _stateService: VideoHubStateService;
  private _loaderService: VideoLoaderService;

  public constructor() {
    super();
    this._stateService = VideoHubStateService.getInstance();
    this._loaderService = VideoLoaderService.getInstance();
  }

  // In VideoGridWebPart.ts

// In VideoGridWebPart.ts
public render(): void {
  // Show loading indicator
  ReactDom.render(
    React.createElement('div', {}, 'Loading videos...'),
    this.domElement
  );

  // Detect if the page is in RTL mode
  const isRTL = document.documentElement.getAttribute('dir') === 'rtl';

  // First, check if the list exists
  console.log('Checking if list exists...');
  this.context.spHttpClient
    .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Title eq '${this.properties.libraryName || 'KMSVideoHub'}'`,
      SPHttpClient.configurations.v1)
    .then(r => r.json())
    .then(result => {
      if (result.value && result.value.length > 0) {
        console.log('List found:', result.value[0].Title);

        // Now try to get a single item to test basic access
        console.log('Testing basic item access...');
        return this.context.spHttpClient
          .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.libraryName || 'KMSVideoHub'}')/items?$top=1`,
            SPHttpClient.configurations.v1)
          .then(r => {
            if (!r.ok) {
              console.error('Error accessing items:', r.status, r.statusText);
              throw new Error(`HTTP ${r.status}`);
            }
            return r.json();
          })
          .then(itemResult => {
            console.log('Basic item access successful, found items:', itemResult.value.length);

// In VideoGridWebPart.ts
// Replace the loadVideosWithThumbnails call with loadVideosWithCreatedAndAuthor
console.log('Loading videos with Created, Author, and metadata...');
const loaderService = VideoLoaderService.getInstance();
return loaderService.loadVideosWithCreatedAndAuthor(
  this.context.pageContext.web.absoluteUrl,
  this.properties.libraryName || 'KMSVideoHub',
  this.context.spHttpClient,
  window.location.href
);
          })
          .then(videos => {
            console.log('Videos loaded successfully:', videos.length);

            // Update the shared state service with the videos
            const stateService = VideoHubStateService.getInstance();
            stateService.setVideos(videos);

            // Render the grid component
            ReactDom.render(
              React.createElement(VideoGrid, {
                isRTL,
                webUrl: this.context.pageContext.web.absoluteUrl
              }),
              this.domElement
            );
          });
      } else {
        console.error('List not found:', this.properties.libraryName || 'KMSVideoHub');
        ReactDom.render(
          React.createElement('p', {}, `⚠️ Library "${this.properties.libraryName || 'KMSVideoHub'}" not found`),
          this.domElement
        );
        return Promise.reject(new Error('List not found'));
      }
    })
    .catch(err => {
      console.error('Error loading videos:', err);
      ReactDom.render(
        React.createElement('p', {}, `⚠️ ${err.message}`),
        this.domElement
      );
    });
}





}
