// //-------------------------------------------------------LEFTy

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

// // //-------------------------------------------------------RIGHTy
