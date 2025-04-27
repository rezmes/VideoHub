// src/webparts/videoGrid/components/VideoGrid.tsx
import * as React from "react";
import {
  DocumentCard,
  DocumentCardTitle,
  DocumentCardPreview,
  IDocumentCardPreviewProps,
} from "office-ui-fabric-react/lib/DocumentCard";
import {
  VideoHubStateService,
  IVideoFilterOptions,
} from "../../../shared/VideoHubStateService";
import { IVideo } from "../../../shared/IVideo";
import styles from "../../videoHub/components/VideoHub.module.scss";

export interface IVideoGridProps {
  isRTL?: boolean;
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

    // Bind methods
    this._handleStateChange = this._handleStateChange.bind(this);
  }

  public componentDidMount(): void {
    // Subscribe to state changes
    this._stateService.subscribe(this._handleStateChange);
    console.log("VideoGrid mounted and subscribed to state changes");
  }

  public componentWillUnmount(): void {
    // Unsubscribe from state changes
    this._stateService.unsubscribe(this._handleStateChange);
    console.log("VideoGrid unmounted and unsubscribed from state changes");
  }

  private _handleStateChange(): void {
    console.log("VideoGrid received state change notification");

    // Update filtered videos when state changes
    this.setState({
      filteredVideos: this._getFilteredVideos(),
      filterOptions: this._stateService.getFilterOptions(),
    });
  }

  private _parseDuration(duration: string): number {
    if (!duration) return 0;

    // Try to parse duration in format "MM:SS" or "HH:MM:SS"
    const parts = duration.split(":").map(function (part) {
      return parseInt(part, 10);
    });

    if (parts.length === 2) {
      // MM:SS format
      return parts[0] * 60 + parts[1];
    } else if (parts.length === 3) {
      // HH:MM:SS format
      return parts[0] * 3600 + parts[1] * 60 + parts[2];
    }

    return 0;
  }

  private _getFilteredVideos(): IVideo[] {
    const videos = this._stateService.getVideos();
    const filterOptions = this._stateService.getFilterOptions();

    console.log("Filtering videos with options:", filterOptions);
    console.log("Total videos before filtering:", videos.length);

    let filtered = videos.slice(); // Create a copy of the array

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
      console.log("After search term filter:", filtered.length);
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
      console.log("After category filter:", filtered.length);
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
      console.log("After department filter:", filtered.length);
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
      console.log("After duration filter:", filtered.length);
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

    console.log("Final filtered videos count:", filtered.length);
    return filtered;
  }

  // src/webparts/videoGrid/components/VideoGrid.tsx
  public render(): React.ReactElement<IVideoGridProps> {
    const { isRTL } = this.props;
    const { filteredVideos } = this.state;
    const cardW = 220;

    // Apply RTL classes conditionally
    const gridClass = isRTL
      ? `${styles.videoGrid} ${styles.videoGridRtl}`
      : styles.videoGrid;

    return (
      <div className={styles.gridContainer}>
        {!filteredVideos || filteredVideos.length === 0 ? (
          <div className={styles.noResults}>
            {this._stateService.getVideos().length === 0
              ? "Loading videos..."
              : "No videos match your search criteria."}
          </div>
        ) : (
          <div className={gridClass}>
            {/* Rest of your render code */}

            {filteredVideos.map(function (v) {
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
                <a key={v.id} href={v.url} className={styles.videoCard}>
                  <div className={styles.cardContent}>
                    <DocumentCard>
                      <DocumentCardPreview {...prev} />
                      <DocumentCardTitle title={v.title} />
                      {v.duration && (
                        <div className={styles.durationText}>
                          Duration: {v.duration}
                        </div>
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
