// src/webparts/videoGrid/components/VideoGrid.tsx
import * as React from "react";
import {
  DocumentCard,
  DocumentCardTitle,
  DocumentCardPreview,
  IDocumentCardPreviewProps,
} from "office-ui-fabric-react/lib/DocumentCard";
import { ImageFit } from "office-ui-fabric-react/lib/Image";
import {
  VideoHubStateService,
  IVideoFilterOptions,
} from "../../../shared/VideoHubStateService";
import { IVideo } from "../../../shared/IVideo";
import styles from './VideoGrid.module.scss';

export interface IVideoGridProps {
  isRTL?: boolean;
  webUrl?: string;
}

export interface IVideoGridState {
  filteredVideos: IVideo[];
  filterOptions: IVideoFilterOptions;
}

export default class VideoGrid extends React.Component<
  IVideoGridProps,
  IVideoGridState
> {
  private _stateService: VideoHubStateService;

  constructor(props: IVideoGridProps) {
    super(props);

    this._stateService = VideoHubStateService.getInstance();

    this.state = {
      filteredVideos: this._getFilteredVideos(),
      filterOptions: this._stateService.getFilterOptions(),
    };

    this._handleStateChange = this._handleStateChange.bind(this);
  }

  public componentDidMount(): void {
    this._stateService.subscribe(this._handleStateChange);
  }

  public componentWillUnmount(): void {
    this._stateService.unsubscribe(this._handleStateChange);
  }

  private _handleStateChange(): void {
    this.setState({
      filteredVideos: this._getFilteredVideos(),
      filterOptions: this._stateService.getFilterOptions(),
    });
  }

// In VideoGrid.tsx
private _getFilteredVideos(): IVideo[] {
  return this._stateService.getFilteredVideos();
}

// Remove the duplicate _parseDuration method and filtering logic


  public render(): React.ReactElement<IVideoGridProps> {
    const { isRTL, webUrl } = this.props;
    const { filteredVideos } = this.state;
    const cardW = 220;

    // Get total video count
    const totalVideos = this._stateService.getVideos().length;

    // Apply RTL classes conditionally
    const gridClass = isRTL
      ? `${styles.videoGrid} ${styles.videoGridRtl}`
      : styles.videoGrid;

    return (
      <div className={styles.gridContainer}>
        <div className={styles.gridHeader}>
          <div className={styles.videoCount}>
            {filteredVideos.length > 0
              ? `نمایش ${filteredVideos.length} از ${totalVideos} ویدئو`
              : totalVideos > 0
              ? "No videos match your search criteria"
              : "Loading videos..."}
          </div>
        </div>

        {!filteredVideos || filteredVideos.length === 0 ? (
          <div className={styles.noResults}>
            {this._stateService.getVideos().length === 0 ? (
              <div className={styles.loadingSpinner}></div>
            ) : (
              "No videos match your search criteria."
            )}
          </div>
        ) : (
          <div className={gridClass}>
            {filteredVideos.map((v) => {
              const prev: IDocumentCardPreviewProps = {
                previewImages: [
                  {
                    previewImageSrc:
                      v.previewUrl ||
                      `${webUrl}/_layouts/15/next/odspnext/odsp-media/images/itemtypes/96/video.png`,
                    width: cardW,
                    height: 125,
                    imageFit: v.previewUrl ? ImageFit.cover : ImageFit.center,
                    iconSrc: !v.previewUrl
                      ? `${webUrl}/_layouts/15/next/odspnext/odsp-media/images/itemtypes/96/video.png`
                      : undefined,
                  },
                ],
              };

              return (
                <a key={v.id} href={v.url} className={styles.videoCard}>
                  <div className={styles.cardContent}>
                    <div className={styles.thumbnailContainer}>
                      {v.previewUrl &&
                      v.previewUrl !== `/_layouts/15/images/videoicon.png` ? (
                        <img
                          src={v.previewUrl}
                          className={styles.thumbnailImage}
                          alt={v.title}
                        />
                      ) : (
                        <div
                          className={styles.defaultThumbnail}
                          style={{
                            backgroundImage: `url('${webUrl}/_layouts/15/next/odspnext/odsp-media/images/itemtypes/96/video.png')`,
                          }}
                        />
                      )}
                      {v.duration && (
                        <div className={styles.durationBadge}>{v.duration}</div>
                      )}
                    </div>
                    <div className={styles.cardBody}>
                      <DocumentCardTitle title={v.title} />
                      <div className={styles.videoStats}>
                        {v.viewCount !== undefined && (
                          <span className={styles.viewCount}>
                            {v.viewCount} {v.viewCount === 1 ? "نمایش" : "نمایش"}
                          </span>
                        )}
                        {v.uploadDate && (
                          <span className={styles.uploadDate}>
                            {v.uploadDate}
                          </span>
                        )}
                      </div>
                      {v.author && (
                        <div className={styles.authorText}>{v.author}</div>
                      )}
                      {v.category && (
                        <div className={styles.metadataText}>
                          دسته: {v.category}
                        </div>
                      )}
                      {v.department && (
                        <div className={styles.metadataText}>
                          Dept: {v.department}
                        </div>
                      )}
                    </div>
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
