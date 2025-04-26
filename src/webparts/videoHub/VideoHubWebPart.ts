// // // //-------------------------------------------------------LEFTy
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox
} from '@microsoft/sp-webpart-base';
import { Version }                            from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as React    from 'react';
import * as ReactDom from 'react-dom';
import VideoHub, { IVideoHubProps } from './components/VideoHub';
import { IVideo }                   from './components/IVideo';

export interface IVideoHubWebPartProps { rtl: boolean; }

export default class VideoHubWebPart
  extends BaseClientSideWebPart<IVideoHubWebPartProps> {

  public render(): void {
    this._fetchVideos()
      .then(videos => {
        const elm = React.createElement<IVideoHubProps>(VideoHub, {
          videos,
          rtl: this.properties.rtl
        });
        ReactDom.render(elm, this.domElement);
      })
      .catch(err => {
        console.error('🎬 VideoHub fetch error:', err);
        ReactDom.render(
          React.createElement('p', null, '⚠️ Unable to load videos.'),
          this.domElement);
      });
  }

  /* ---------- REST ---------- */
  private _fetchVideos(): Promise<IVideo[]> {

    const web = this.context.pageContext.web.absoluteUrl;

    const url =
      `${web}/_api/web/lists/getbytitle('Videos')/items` +
      `?$select=Id,Title,EncodedAbsUrl,EncodedAbsThumbnailUrl` +
      `&$filter=File_x0020_Type eq 'mp4'` +
      `&$orderby=Id desc&$top=30`;

    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => {
        if (!r.ok) { throw new Error('HTTP ' + r.status); }
        return r.json();
      })
      .then(j => (j.value || []).map(it => ({
        id:    it.Id,
        title: it.Title || (it.EncodedAbsUrl as string).split('/').pop(),
        url:   it.EncodedAbsUrl,
        previewUrl: it.EncodedAbsThumbnailUrl || ''
      } as IVideo)));
  }

  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'تنظیمات' },
        groups: [{
          groupFields: [
            PropertyPaneCheckbox('rtl', { text: 'Right-to-Left layout' })
          ]
        }]
      }]
    };
  }
}
// // //-------------------------------------------------------RIGHTy
