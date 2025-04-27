// // //-------------------------------------------------------LEFTy
// VideoHub.tsx - Updated with filtering and sorting
import * as React from "react";
import { IVideo } from "./IVideo";
import { IVideoHubState } from "./VideoHubState";
import {
  DocumentCard,
  DocumentCardTitle,
  DocumentCardPreview,
  IDocumentCardPreviewProps,
} from "office-ui-fabric-react/lib/DocumentCard";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { DefaultButton } from "office-ui-fabric-react/lib/Button";

export interface IVideoHubProps {
  videos: IVideo[];
}

export default class VideoHub extends React.Component<
  IVideoHubProps,
  IVideoHubState
> {
  constructor(props: IVideoHubProps) {
    super(props);
    this.state = {
      filteredVideos: props.videos,
      searchTerm: "",
      sortBy: "title",
      sortDirection: "asc",
    };
  }

  public componentWillReceiveProps(nextProps: IVideoHubProps): void {
    if (nextProps.videos !== this.props.videos) {
      this.setState({ filteredVideos: nextProps.videos }, () => {
        this.filterAndSortVideos();
      });
    }
  }

  private handleSearch = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this.setState({ searchTerm: newValue || "" }, () => {
      this.filterAndSortVideos();
    });
  };

  private handleSortChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState({ sortBy: option.key as string }, () => {
        this.filterAndSortVideos();
      });
    }
  };

  private toggleSortDirection = (): void => {
    this.setState(
      {
        sortDirection: this.state.sortDirection === "asc" ? "desc" : "asc",
      },
      () => {
        this.filterAndSortVideos();
      }
    );
  };

  private filterAndSortVideos = (): void => {
    const { searchTerm, sortBy, sortDirection } = this.state;
    let filtered = [...this.props.videos];

    // Filter by search term
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      filtered = filtered.filter(
        (video) => video.title.toLowerCase().indexOf(term) !== -1
      );
    }

    // Sort videos
    filtered.sort((a, b) => {
      let comparison = 0;

      if (sortBy === "title") {
        comparison = a.title.localeCompare(b.title);
      } else if (sortBy === "duration") {
        // Simple string comparison for duration
        comparison = (a.duration || "").localeCompare(b.duration || "");
      }

      return sortDirection === "asc" ? comparison : -comparison;
    });

    this.setState({ filteredVideos: filtered });
  };

  public render(): React.ReactElement<IVideoHubProps> {
    const { filteredVideos, sortDirection } = this.state;

    const containerStyle: React.CSSProperties = {
      margin: "0 -8px",
    };

    const controlsStyle: React.CSSProperties = {
      display: "flex",
      marginBottom: 16,
      flexWrap: "wrap",
      gap: 8,
    };

    const wrap: React.CSSProperties = {
      display: "flex",
      flexWrap: "wrap",
      gap: 12,
    };

    const cardW = 220;

    const durationStyle: React.CSSProperties = {
      padding: "0 16px 8px 16px",
      fontSize: "12px",
      color: "#666666",
    };

    const sortOptions: IDropdownOption[] = [
      { key: "title", text: "Title" },
      { key: "duration", text: "Duration" },
    ];

    return (
      <div style={containerStyle}>
        <div style={controlsStyle}>
          <TextField
            placeholder="Search videos..."
            onChange={this.handleSearch}
            style={{ width: 200, marginRight: 8 }}
          />
          <Dropdown
            label="Sort by:"
            selectedKey={this.state.sortBy}
            options={sortOptions}
            onChange={this.handleSortChange}
            style={{ width: 120, marginRight: 8 }}
          />
          <DefaultButton
            text={`Order: ${sortDirection === "asc" ? "A to Z" : "Z to A"}`}
            onClick={this.toggleSortDirection}
          />
        </div>

        {filteredVideos.length === 0 ? (
          <div>No videos match your search criteria.</div>
        ) : (
          <div style={wrap}>
            {filteredVideos.map((v) => {
              const prev: IDocumentCardPreviewProps = {
                previewImages: [
                  {
                    previewImageSrc:
                      v.previewUrl || `/_layouts/15/images/videoicon.png`,
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
                      {v.duration && (
                        <div style={durationStyle}>Duration: {v.duration}</div>
                      )}
                    </DocumentCard>
                  </div>
                </a>
              );
            })}
          </div>
        )}
      </div>
    );
  }
}

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
//     const durationStyle: React.CSSProperties = {
//       padding: "0 16px 8px 16px",
//       fontSize: "12px",
//       color: "#666666",
//     };

//     return (
//       <div style={wrap}>
//         {this.props.videos.map((v) => {
//           const prev: IDocumentCardPreviewProps = {
//             previewImages: [
//               {
//                 previewImageSrc:
//                   v.previewUrl || `/_layouts/15/images/videoicon.png`,
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
//                   {v.duration && (
//                     <div style={durationStyle}>Duration: {v.duration}</div>
//                   )}
//                 </DocumentCard>
//               </div>
//             </a>
//           );
//         })}
//       </div>
//     );
//   }
// }
