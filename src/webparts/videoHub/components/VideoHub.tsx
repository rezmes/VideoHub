// //-------------------------------------------------------RIGHTy

import * as React from "react";
import { IVideo } from "./IVideo";
import {
  DocumentCard,
  DocumentCardTitle,
  DocumentCardPreview,
  IDocumentCardPreviewProps,
} from "office-ui-fabric-react/lib/DocumentCard";

export interface IVideoHubProps {
  videos: IVideo[];
  rtl: boolean;
}

export default class VideoHub extends React.Component<IVideoHubProps, {}> {
  public render(): React.ReactElement<IVideoHubProps> {
    const wrap: React.CSSProperties = {
      display: "flex",
      flexWrap: "wrap",
      gap: 12,
      direction: this.props.rtl ? "rtl" : "ltr",
      justifyContent: this.props.rtl ? "flex-end" : "flex-start",
    };

    const cardWidth = 220;

    return (
      <div style={wrap}>
        {this.props.videos.map((v) => {
          const prev: IDocumentCardPreviewProps = {
            previewImages: [
              {
                previewImageSrc:
                  v.previewUrl || "/_layouts/15/images/video.png",
                width: cardWidth,
                height: 125,
              },
            ],
          };

          return (
            <a key={v.id} href={v.url} style={{ textDecoration: "none" }}>
              <div style={{ width: cardWidth }}>
                <DocumentCard>
                  <DocumentCardPreview {...prev} />
                  <DocumentCardTitle title={v.title} />
                </DocumentCard>
              </div>
            </a>
          );
        })}
      </div>
    );
  }
}

// // //-------------------------------------------------------LEFTy
// import * as React from "react";
// import { IVideo } from "./IVideo";
// import {
//   DocumentCard,
//   DocumentCardTitle,
//   DocumentCardPreview,
//   IDocumentCardPreviewProps,
// } from "office-ui-fabric-react/lib/DocumentCard";

// export interface IVideoHubProps {
//   videos: IVideo[];
// }

// export default class VideoHub extends React.Component<IVideoHubProps, {}> {
//   public render(): React.ReactElement<IVideoHubProps> {
//     const wrap: React.CSSProperties = {
//       display: "flex",
//       flexWrap: "wrap",
//       gap: 12,
//     };

//     const cardW = 220;

//     return (
//       <div style={wrap}>
//         {this.props.videos.map((v) => {
//           const prev: IDocumentCardPreviewProps = {
//             previewImages: [
//               {
//                 previewImageSrc:
//                   v.previewUrl || "/_layouts/15/images/icvideo.png",
//                 width: cardW,
//                 height: 125,
//               },
//             ],
//           };

//           return (
//             <a key={v.id} href={v.url} style={{ textDecoration: "none" }}>
//               <div style={{ width: cardW }}>
//                 <DocumentCard>
//                   <DocumentCardPreview {...prev} />
//                   <DocumentCardTitle title={v.title} />
//                 </DocumentCard>
//               </div>
//             </a>
//           );
//         })}
//       </div>
//     );
//   }
// }

// import * as React from 'react';
// import styles from './VideoHub.module.scss';
// import { IVideoHubProps } from './IVideoHubProps';
// import { escape } from '@microsoft/sp-lodash-subset';

// export default class VideoHub extends React.Component < IVideoHubProps, {} > {
//   public render(): React.ReactElement<IVideoHubProps> {
//     return(
//       <div className = { styles.videoHub } >
//   <div className={styles.container}>
//     <div className={styles.row}>
//       <div className={styles.column}>
//         <span className={styles.title}>Welcome to SharePoint!</span>
//         <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
//         <p className={styles.description}>{escape(this.props.description)}</p>
//         <a href='https://aka.ms/spfx' className={styles.button}>
//           <span className={styles.label}>Learn more</span>
//         </a>
//       </div>
//     </div>
//   </div>
//       </div >
//     );
//   }
// }
