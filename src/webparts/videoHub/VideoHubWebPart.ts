// // // // //-------------------------------------------------------LEFTy
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as React    from 'react';
import * as ReactDom from 'react-dom';

import VideoHub   from './components/VideoHub';
import { IVideo } from './components/IVideo';

export interface IVideoHubWebPartProps { libraryName: string; }

export default class VideoHubWebPart
  extends BaseClientSideWebPart<IVideoHubWebPartProps> {

  public render(): void {
    this._load()
      .then(videos =>
        ReactDom.render(React.createElement(VideoHub,{ videos }), this.domElement)
      )
      .catch(err =>
        ReactDom.render(React.createElement('p',{}, `⚠️ ${err.message}`), this.domElement)
      );
  }

  /* ---------- REST + thumbnail picker ---------- */
  private _load(): Promise<IVideo[]> {

    const list = this.properties.libraryName || 'KMSVideos';

    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('${list}')/items` +
      `?$select=Id,Title,EncodedAbsUrl,ThumbnailURL,url,FileLeafRef,File/ServerRelativeUrl` +
      `&$expand=File` +
      `&$filter=File_x0020_Type eq 'mp4'` +
      `&$orderby=Modified desc&$top=30`;

    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => r.json())
      .then(j => (j.value || []).map((it:any) => {

        /* choose thumbnail */
        let thumb: string = '';

        if (it.ThumbnailURL && it.ThumbnailURL.Url) {
          thumb = it.ThumbnailURL.Url;
        } else if (it.url && it.url.Url) {
          thumb = it.url.Url;
        } else {
          /* build “Preview Images/<name>_mp4.png” */
          const folder = it.EncodedAbsUrl.substring(0, it.EncodedAbsUrl.lastIndexOf('/'));
          const bare   = encodeURIComponent(
                           decodeURIComponent(it.FileLeafRef).replace(/\.[^/.]+$/, ''));
          thumb = `${folder}/Preview%20Images/${bare}_mp4.png`;
        }

        /* fallback icon if still empty */
        if (!thumb) {
          thumb = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/images/icvideo.png`;
        }

        /* clean title */
        const tit = decodeURIComponent(it.Title || it.FileLeafRef || 'Video');

        return {
          id: it.Id,
          title: tit,
          url:  it.EncodedAbsUrl,
          previewUrl: thumb
        } as IVideo;
      }));
  }

  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages:[{
        header:{ description:'Settings' },
        groups:[{
          groupFields:[
            PropertyPaneTextField('libraryName',{ label:'Library title', value:'KMSVideos' })
          ]
        }]
      }]
    };
  }
}






















// // //-------------------------------------------------------RIGHTy
// VideoHubWebPart.ts
// import {
//   BaseClientSideWebPart
// } from '@microsoft/sp-webpart-base';
// import { Version } from '@microsoft/sp-core-library';
// import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
// import * as React from 'react';
// import * as ReactDom from 'react-dom';
// import {
//   IPropertyPaneConfiguration,
//   PropertyPaneTextField
// } from '@microsoft/sp-webpart-base';
// import VideoHub from './components/VideoHub';
// import { IVideo } from './components/IVideo';

// export interface IVideoHubWebPartProps {
//   libraryName: string;
// }

// export default class VideoHubWebPart
//   extends BaseClientSideWebPart<IVideoHubWebPartProps> {

//   public render(): void {
//     this._getVideos()
//       .then(videos => {
//         ReactDom.render(React.createElement(VideoHub, { videos }), this.domElement);
//       })
//       .catch(err => {
//         ReactDom.render(
//           React.createElement('div', {}, [
//             React.createElement('h3', {}, '⚠️ Error loading videos'),
//             React.createElement('p', {}, err.message)
//           ]),
//           this.domElement
//         );
//       });
//   }

//   private _getVideos(): Promise<IVideo[]> {
//     const libraryName = this.properties.libraryName || 'KMSVideos';
//     const url =
//       `${this.context.pageContext.web.absoluteUrl}` +
//       `/_api/web/lists/getbytitle('${libraryName}')/items` +
//       `?$select=Id,Title,EncodedAbsUrl,ThumbnailURL,url,File_x0020_Type,FileLeafRef` +
//       `&$filter=File_x0020_Type eq 'mp4'` +
//       `&$orderby=Modified desc&$top=30`;

//     return this.context.spHttpClient
//       .get(url, SPHttpClient.configurations.v1)
//       .then((r: SPHttpClientResponse) => {
//         if (!r.ok) { throw new Error(`HTTP ${r.status}`); }
//         return r.json();
//       })
//       .then(data => this._processVideoItems(data.value || []));
//   }

//   private _processVideoItems(items: any[]): IVideo[] {
//     if (!items || items.length === 0) {
//       return [];
//     }

//     return items.map((item: any) => {
//       // Try to get the thumbnail URL from various possible sources
//       let thumbnailUrl = '';

//       // First check if ThumbnailURL exists and has a Url property
//       if (item.ThumbnailURL && typeof item.ThumbnailURL === 'object' && item.ThumbnailURL.Url) {
//         thumbnailUrl = item.ThumbnailURL.Url;
//       }
//       // Check if url field exists and has a Url property
//       else if (item.url && typeof item.url === 'object' && item.url.Url) {
//         thumbnailUrl = item.url.Url;
//       }
//       // For the "simens" video specifically - hardcode the known working path
//       else if (item.FileLeafRef && item.FileLeafRef.toLowerCase().includes('simens')) {
//         thumbnailUrl = `${this.context.pageContext.web.absoluteUrl}/sites/kms/KMSVideos/simens/Preview%20Images/simens.png`;
//       }

//       // If we still don't have a thumbnail, use a generic icon
//       if (!thumbnailUrl) {
//         thumbnailUrl = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/images/videoicon.png`;
//       }

//       // Clean up the title
//       let title = item.Title || '';
//       if (!title && item.FileLeafRef) {
//         title = item.FileLeafRef.replace(/\.[^/.]+$/, ""); // Remove extension
//         title = decodeURIComponent(title); // Decode URL encoding
//       }

//       return {
//         id: item.Id,
//         title: title || 'Untitled Video',
//         url: item.EncodedAbsUrl || '',
//         previewUrl: thumbnailUrl
//       } as IVideo;
//     });
//   }

//   protected onDispose(): void {
//     ReactDom.unmountComponentAtNode(this.domElement);
//   }

//   protected get dataVersion(): Version {
//     return Version.parse('1.0');
//   }

//   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
//     return {
//       pages: [{
//         header: { description: 'Settings' },
//         groups: [{
//           groupFields: [
//             PropertyPaneTextField('libraryName', {
//               label: 'Library title',
//               value: 'KMSVideos'
//             })
//           ]
//         }]
//       }]
//     };
//   }
// }

