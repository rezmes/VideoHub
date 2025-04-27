// src/webparts/videoGrid/VideoGridWebPart.ts
// import * as strings from '../videoHub/loc/mystrings';
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
// src/webparts/videoGrid/VideoGridWebPart.ts
public render(): void {
  // Show loading indicator
  ReactDom.render(
    React.createElement('div', {}, 'Loading videos...'),
    this.domElement
  );

  // Detect if the page is in RTL mode
  const isRTL = document.documentElement.getAttribute('dir') === 'rtl';

  // Get the state service instance
  const stateService = VideoHubStateService.getInstance();
  console.log('Grid web part using state service instance:', stateService);

  // Render the grid component first so it can subscribe to state changes
  ReactDom.render(
    React.createElement(VideoGrid, { isRTL }),
    this.domElement
  );

  // Then load videos
  const list = this.properties.libraryName || 'KMSVideoHub';
  const currentPageUrl = window.location.href;

  this._loaderService.loadVideos(
    this.context.pageContext.web.absoluteUrl,
    list,
    this.context.spHttpClient,
    currentPageUrl
  )
    .then(videos => {
      console.log('Videos loaded successfully:', videos.length);

      // Update the shared state service with the videos
      stateService.setVideos(videos);
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
