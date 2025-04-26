// // //-------------------------------------------------------LEFTy

// import * as React from 'react';
// import { IVideo } from './IVideo';
// import {
//   DocumentCard, DocumentCardPreview, DocumentCardTitle,
//   IDocumentCardPreviewProps
// } from 'office-ui-fabric-react/lib/DocumentCard';
// import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';

// export interface IVideoHubProps { videos: IVideo[]; }
// export interface IVideoHubState { play?: string; }

// export default class VideoHub
//   extends React.Component<IVideoHubProps, IVideoHubState> {

//   public state: IVideoHubState = { play: undefined };

//   /* open / close helpers */
//   private _open = (u: string) => this.setState({ play: u });
//   private _close = ()          => this.setState({ play: undefined });

//   public render(): React.ReactElement<IVideoHubProps> {

//     const wrapper: React.CSSProperties = { display:'flex', flexWrap:'wrap', gap:12 };
//     const cardW = 220;
//     const durCss: React.CSSProperties = { padding:'0 16px 8px', fontSize:12, color:'#666' };

//     return (
//       <div>
//         <div style={wrapper}>
//           {this.props.videos.map(v => {
//             const preview: IDocumentCardPreviewProps = {
//               previewImages: [{
//                 previewImageSrc: v.previewUrl || '/_layouts/15/images/videoicon.png',
//                 width: cardW, height:125
//               }]
//             };
//             return (
//               <div key={v.id} style={{ width:cardW }}>
//                 <DocumentCard onClick={() => this._open(v.url)}>
//                   <DocumentCardPreview {...preview}/>
//                   <DocumentCardTitle title={v.title}/>
//                   {v.duration &&
//                     <div style={durCss}>Duration: {v.duration}</div>}
//                 </DocumentCard>
//               </div>
//             );
//           })}
//         </div>

//         {this.state.play &&
//           <Dialog
//             hidden={false}
//             dialogContentProps={{ type: DialogType.largeHeader }}
//             modalProps={{ isBlocking:false }}
//             onDismiss={this._close}>
//             <video src={this.state.play} controls style={{ width:'100%' }}/>
//           </Dialog>}
//       </div>
//     );
//   }
// }
// // //-------------------------------------------------------RIGHTy


import * as React from 'react';
import { IVideo } from './IVideo';
import {
  DocumentCard,
  DocumentCardTitle,
  DocumentCardPreview,
  IDocumentCardPreviewProps
} from 'office-ui-fabric-react/lib/DocumentCard';

export interface IVideoHubProps { videos: IVideo[]; }

export default class VideoHub extends React.Component<IVideoHubProps, {}> {

  public render(): React.ReactElement<IVideoHubProps> {

    const wrap: React.CSSProperties = { display:'flex', flexWrap:'wrap', gap:12 };
    const cardW = 220;
    const durationStyle: React.CSSProperties = { 
      padding: '0 16px 8px 16px',
      fontSize: '12px',
      color: '#666666'
    };

    return (
      <div style={wrap}>
        {this.props.videos.map(v => {
          const prev: IDocumentCardPreviewProps = {
            previewImages: [{
              previewImageSrc: v.previewUrl || `/_layouts/15/images/videoicon.png`,
              width: cardW,
              height: 125
            }]
          };
          
          return (
            <a key={v.id} href={v.url} style={{ textDecoration:'none' }}>
              <div style={{ width: cardW }}>
                <DocumentCard>
                  <DocumentCardPreview {...prev}/>
                  <DocumentCardTitle title={v.title} />
                  {v.duration && 
                    <div style={durationStyle}>Duration: {v.duration}</div>
                  }
                </DocumentCard>
              </div>
            </a>
          );
        })}
      </div>
    );
  }
}