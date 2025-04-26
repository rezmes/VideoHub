// // // // //-------------------------------------------------------LEFTy
// import {
//   BaseClientSideWebPart,
//   IPropertyPaneConfiguration,
//   PropertyPaneTextField
// } from '@microsoft/sp-webpart-base';
// import { Version } from '@microsoft/sp-core-library';
// import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
// import * as React    from 'react';
// import * as ReactDom from 'react-dom';

// import VideoHub from './components/VideoHub';
// import { IVideo }  from './components/IVideo';

// export interface IVideoHubWebPartProps { libraryName: string; }

// export default class VideoHubWebPart
//   extends BaseClientSideWebPart<IVideoHubWebPartProps> {

//   public render(): void {
//     this._load()
//       .then(v => ReactDom.render(React.createElement(VideoHub,{ videos:v }), this.domElement))
//       .catch(e => ReactDom.render(React.createElement('p',{},'⚠ '+e.message), this.domElement));
//   }

//   /* -------- REST ---------- */
//   private _load(): Promise<IVideo[]> {

//     const list = this.properties.libraryName || 'KMSVideoHub';

//     const url =
//       `${this.context.pageContext.web.absoluteUrl}` +
//       `/_api/web/lists/getbytitle('${list}')/items` +
//       `?$select=Id,Title,FileRef,FileLeafRef,ThumbnailURL,MediaDuration` +
//       `&$filter=File_x0020_Type eq 'mp4'` +
//       `&$orderby=Modified desc&$top=30`;

//     return this.context.spHttpClient
//       .get(url, SPHttpClient.configurations.v1)
//       .then((r:SPHttpClientResponse)=>{ if(!r.ok){throw new Error('HTTP '+r.status);} return r.json();})
//       .then(j => (j.value as any[] || []).map(it => {

//         const thumb = (it.ThumbnailURL && it.ThumbnailURL.Url)
//             ? it.ThumbnailURL.Url
//             : `${this.context.pageContext.web.absoluteUrl}/_layouts/15/images/videoicon.png`;

//         /* modern player url */
//         const fileRef = it.FileRef as string;
//         const encodedFile = encodeURIComponent(fileRef);
//         const parent = fileRef.substring(0, fileRef.lastIndexOf('/'));
//         const encodedParent = encodeURIComponent(parent);
//         const modernUrl = `${parent}/Forms/AllItems.aspx?id=${encodedFile}&parent=${encodedParent}`;

//         return {
//           id: it.Id,
//           title: decodeURIComponent(it.Title || it.FileLeafRef || 'Video'),
//           url: modernUrl,
//           previewUrl: thumb,
//           duration: it.MediaDuration || ''
//         } as IVideo;
//       }));
//   }

//   protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
//   protected get dataVersion(): Version { return Version.parse('1.0'); }

//   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
//     return {
//       pages:[{
//         header:{ description:'Settings' },
//         groups:[{
//           groupFields:[
//             PropertyPaneTextField('libraryName',{ label:'Library title', value:'KMSVideoHub' })
//           ]
//         }]
//       }]
//     };
//   }
// }






// // //-------------------------------------------------------RIGHTy

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import VideoHub from './components/VideoHub';
import { IVideo } from './components/IVideo';

export interface IVideoHubWebPartProps { libraryName: string; }

export default class VideoHubWebPart
  extends BaseClientSideWebPart<IVideoHubWebPartProps> {

  public render(): void {
    this._load()
      .then(videos =>
        ReactDom.render(React.createElement(VideoHub, { videos }), this.domElement)
      )
      .catch(err =>
        ReactDom.render(React.createElement('p', {}, `⚠️ ${err.message}`), this.domElement)
      );
  }

  /* ---------- REST + thumbnail picker ---------- */
  private _load(): Promise<IVideo[]> {

    const list = this.properties.libraryName || 'KMSVideoHub';
    
    // Get the current page URL for the return link
    const currentPageUrl = window.location.href;
    
    // Simplified query - no filter, we'll filter client-side
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
        const mp4Items = items.filter(item => 
          item.FileLeafRef && item.FileLeafRef.toLowerCase().endsWith('.mp4')
        );
        
        return mp4Items.map((it: any) => {
          // Get thumbnail URL from ThumbnailURL field
          const thumb = it.ThumbnailURL && it.ThumbnailURL.Url 
            ? it.ThumbnailURL.Url 
            : `${this.context.pageContext.web.absoluteUrl}/_layouts/15/images/videoicon.png`;
          
          // Create the modern player URL based on the example
          const fileRef = it.FileRef || '';
          
          // Format: /sites/kms/KMSVideoHub/Forms/AllItems.aspx?id=%2Fsites%2Fkms%2FKMSVideoHub%2FFileName&parent=%2Fsites%2Fkms%2FKMSVideoHub&source=<currentPageUrl>
          const encodedFileRef = encodeURIComponent(fileRef);
          const parentFolder = fileRef.substring(0, fileRef.lastIndexOf('/'));
          const encodedParentFolder = encodeURIComponent(parentFolder);
          const encodedSourceUrl = encodeURIComponent(currentPageUrl);
          
          // Use 'source' parameter instead of 'src' for return URL
//          // Build modern viewer URL that returns to the gallery
// const modernPlayerUrl =
// `${parentFolder}/Forms/AllItems.aspx?Source=${encodeURIComponent(currentPageUrl)}` +
// `&id=${encodedFileRef}&parent=${encodedParentFolder}`;
// Build modern viewer URL that returns to the gallery
const modernPlayerUrl =
  `${parentFolder}/Forms/AllItems.aspx?Source=${encodeURIComponent(currentPageUrl)}` +
  `&id=${encodedFileRef}&parent=${encodedParentFolder}`;
            
          return {
            id: it.Id,
            title: it.Title || (it.FileLeafRef ? it.FileLeafRef.replace('.mp4', '') : 'Untitled Video'),
            url: modernPlayerUrl,
            previewUrl: thumb,
            duration: it.MediaDuration || ''
          } as IVideo;
        });
      });
  }

  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }

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