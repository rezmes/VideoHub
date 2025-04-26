// // //-------------------------------------------------------LEFTy

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
}

export default class VideoHub extends React.Component<IVideoHubProps, {}> {
  public render(): React.ReactElement<IVideoHubProps> {
    const wrap: React.CSSProperties = {
      display: "flex",
      flexWrap: "wrap",
      gap: 12,
    };
    const cardW = 220;

    return (
      <div style={wrap}>
        {this.props.videos.map((v) => {
          const prev: IDocumentCardPreviewProps = {
            previewImages: [
              {
                previewImageSrc: v.previewUrl,
                width: cardW,
                height: 125,
              },
            ],
          };
          return (
            <a key={v.id} href={v.url} style={{ textDecoration: "none" }}>
              <div style={{ width: cardW }}>
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

// // //-------------------------------------------------------RIGHTy
// VideoHub.tsx
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

//     const cardWidth = 220;

//     return (
//       <div style={wrap}>
//         {this.props.videos && this.props.videos.length > 0 ? (
//           this.props.videos.map((v) => {
//             const prev: IDocumentCardPreviewProps = {
//               previewImages: [
//                 {
//                   previewImageSrc:
//                     v.previewUrl ||
//                     "/sites/kms/_layouts/15/images/videoicon.png",
//                   width: cardWidth,
//                   height: 125,
//                 },
//               ],
//             };

//             return (
//               <a key={v.id} href={v.url} style={{ textDecoration: "none" }}>
//                 <div style={{ width: cardWidth }}>
//                   <DocumentCard>
//                     <DocumentCardPreview {...prev} />
//                     <DocumentCardTitle title={v.title} />
//                   </DocumentCard>
//                 </div>
//               </a>
//             );
//           })
//         ) : (
//           <div>No videos found</div>
//         )}
//       </div>
//     );
//   }
// }
