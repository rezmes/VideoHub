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
import styles from "../../videoHub/components/VideoHub.module.scss";

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

  private _parseDuration(duration: string): number {
    if (!duration) return 0;

    const parts = duration.split(":").map(function (part) {
      return parseInt(part, 10);
    });

    if (parts.length === 2) {
      return parts[0] * 60 + parts[1];
    } else if (parts.length === 3) {
      return parts[0] * 3600 + parts[1] * 60 + parts[2];
    }

    return 0;
  }

  private _getFilteredVideos(): IVideo[] {
    const videos = this._stateService.getVideos();
    const filterOptions = this._stateService.getFilterOptions();

    let filtered = videos.slice();

    // Filter by search term
    if (filterOptions.searchTerm) {
      const term = filterOptions.searchTerm.toLowerCase();
      const tempFiltered = [];
      for (let i = 0; i < filtered.length; i++) {
        if (filtered[i].title.toLowerCase().indexOf(term) !== -1) {
          tempFiltered.push(filtered[i]);
        }
      }
      filtered = tempFiltered;
    }

    // Filter by category
    if (filterOptions.selectedCategory) {
      const tempFiltered = [];
      for (let i = 0; i < filtered.length; i++) {
        if (filtered[i].category === filterOptions.selectedCategory) {
          tempFiltered.push(filtered[i]);
        }
      }
      filtered = tempFiltered;
    }

    // Filter by department
    if (filterOptions.selectedDepartment) {
      const tempFiltered = [];
      for (let i = 0; i < filtered.length; i++) {
        if (filtered[i].department === filterOptions.selectedDepartment) {
          tempFiltered.push(filtered[i]);
        }
      }
      filtered = tempFiltered;
    }

    // Filter by duration
    if (filterOptions.durationFilter !== "all") {
      const tempFiltered = [];
      for (let i = 0; i < filtered.length; i++) {
        const durationInSeconds = this._parseDuration(filtered[i].duration);

        if (
          filterOptions.durationFilter === "short" &&
          durationInSeconds < 300
        ) {
          tempFiltered.push(filtered[i]);
        } else if (
          filterOptions.durationFilter === "medium" &&
          durationInSeconds >= 300 &&
          durationInSeconds < 900
        ) {
          tempFiltered.push(filtered[i]);
        } else if (
          filterOptions.durationFilter === "long" &&
          durationInSeconds >= 900
        ) {
          tempFiltered.push(filtered[i]);
        }
      }
      filtered = tempFiltered;
    }

    // Sort videos
    filtered.sort((a, b) => {
      let comparison = 0;

      if (filterOptions.sortBy === "title") {
        comparison = a.title.localeCompare(b.title);
      } else if (filterOptions.sortBy === "duration") {
        comparison =
          this._parseDuration(a.duration) - this._parseDuration(b.duration);
      } else if (filterOptions.sortBy === "category") {
        comparison = (a.category || "").localeCompare(b.category || "");
      } else if (filterOptions.sortBy === "department") {
        comparison = (a.department || "").localeCompare(b.department || "");
      }

      return filterOptions.sortDirection === "asc" ? comparison : -comparison;
    });

    return filtered;
  }

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
              ? `Showing ${filteredVideos.length} of ${totalVideos} videos`
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
                            {v.viewCount} {v.viewCount === 1 ? "view" : "views"}
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
                          Category: {v.category}
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
