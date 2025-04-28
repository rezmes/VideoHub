// src/webparts/videoFilter/VideoFilterWebPart.ts
// import * as strings from '../videoHub/loc/mystrings';

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import VideoFilter from './components/VideoFilter';
import { VideoHubStateService } from '../../shared/VideoHubStateService';

export interface IVideoFilterWebPartProps {
  libraryName: string;
}

export default class VideoFilterWebPart extends BaseClientSideWebPart<IVideoFilterWebPartProps> {
  private _stateService: VideoHubStateService;

  public constructor() {
    super();
    this._stateService = VideoHubStateService.getInstance();
  }

  public render(): void {
    // Detect if the page is in RTL mode
    const isRTL = document.documentElement.getAttribute('dir') === 'rtl';

    ReactDom.render(
      React.createElement(VideoFilter, {
        isRTL
      }),
      this.domElement
    );
  // Add this to both web parts' render methods to verify they're using the same instance
console.log('State service instance ID:', this._stateService);

}

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: "Configure the Filter Video web part" },
        groups: [{
          groupName: "Settings",
          groupFields: [
            PropertyPaneTextField('libraryName', {
              label: "Library Name",
              value: 'KMSVideoHub'
            })
          ]
        }]
      }]
    };
  }
}