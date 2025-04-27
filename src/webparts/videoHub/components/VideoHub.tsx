// // //-------------------------------------------------------LEFTy
// VideoHub.tsx - Updated with filtering and sorting
// VideoHub.tsx
import * as React from "react";
import { IVideo } from "../../../shared/IVideo";
import {
  DocumentCard,
  DocumentCardTitle,
  DocumentCardPreview,
  IDocumentCardPreviewProps,
} from "office-ui-fabric-react/lib/DocumentCard";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { DefaultButton } from "office-ui-fabric-react/lib/Button";
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from "office-ui-fabric-react/lib/ChoiceGroup";
import styles from "./VideoHub.module.scss";

export interface IVideoHubProps {
  videos: IVideo[];
  isRTL?: boolean;
}

export interface IVideoHubState {
  filteredVideos: IVideo[];
  searchTerm: string;
  selectedCategory: string;
  selectedDepartment: string;
  durationFilter: string;
  sortBy: string;
  sortDirection: "asc" | "desc";
  categoryOptions: IDropdownOption[];
  departmentOptions: IDropdownOption[];
}

export default class VideoHub extends React.Component<
  IVideoHubProps,
  IVideoHubState
> {
  constructor(props: IVideoHubProps) {
    super(props);

    // Extract unique categories and departments for dropdown options
    const categoryOptions = this._getUniqueOptions(props.videos, "category");
    const departmentOptions = this._getUniqueOptions(
      props.videos,
      "department"
    );

    this.state = {
      filteredVideos: props.videos,
      searchTerm: "",
      selectedCategory: "",
      selectedDepartment: "",
      durationFilter: "all",
      sortBy: "title",
      sortDirection: "asc",
      categoryOptions,
      departmentOptions,
    };
  }

  private _getUniqueOptions(
    videos: IVideo[],
    field: string
  ): IDropdownOption[] {
    // Create a map of unique values (compatible with ES5)
    const uniqueValues = {};

    videos.forEach((video) => {
      const value = video[field];
      if (value) {
        uniqueValues[value] = true;
      }
    });

    // Convert to array of options (compatible with ES5)
    const options: IDropdownOption[] = [];
    for (const key in uniqueValues) {
      if (uniqueValues.hasOwnProperty(key)) {
        options.push({
          key: key,
          text: key,
        });
      }
    }

    // Create result array with "All" option first
    const result: IDropdownOption[] = [{ key: "", text: "All" }];

    // Add each option individually to avoid type issues
    for (let i = 0; i < options.length; i++) {
      result.push(options[i]);
    }

    return result;
  }

  public componentWillReceiveProps(nextProps: IVideoHubProps): void {
    if (nextProps.videos !== this.props.videos) {
      // Update dropdown options when videos change
      const categoryOptions = this._getUniqueOptions(
        nextProps.videos,
        "category"
      );
      const departmentOptions = this._getUniqueOptions(
        nextProps.videos,
        "department"
      );

      this.setState(
        {
          filteredVideos: nextProps.videos,
          categoryOptions,
          departmentOptions,
        },
        () => {
          this._filterAndSortVideos();
        }
      );
    }
  }

  private _handleSearch = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this.setState({ searchTerm: newValue || "" }, () => {
      this._filterAndSortVideos();
    });
  };

  private _handleCategoryChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState({ selectedCategory: option.key as string }, () => {
        this._filterAndSortVideos();
      });
    }
  };

  private _handleDepartmentChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState({ selectedDepartment: option.key as string }, () => {
        this._filterAndSortVideos();
      });
    }
  };

  private _handleDurationFilterChange = (
    ev?: React.FormEvent<HTMLElement | HTMLInputElement>,
    option?: IChoiceGroupOption
  ): void => {
    if (option) {
      this.setState({ durationFilter: option.key }, () => {
        this._filterAndSortVideos();
      });
    }
  };

  private _handleSortChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState({ sortBy: option.key as string }, () => {
        this._filterAndSortVideos();
      });
    }
  };

  private _toggleSortDirection = (): void => {
    this.setState(
      {
        sortDirection: this.state.sortDirection === "asc" ? "desc" : "asc",
      },
      () => {
        this._filterAndSortVideos();
      }
    );
  };

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

  private _filterAndSortVideos = (): void => {
    const {
      searchTerm,
      selectedCategory,
      selectedDepartment,
      durationFilter,
      sortBy,
      sortDirection,
    } = this.state;

    let filtered = this.props.videos.slice(); // Create a copy of the array

    // Filter by search term
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      filtered = filtered.filter(function (video) {
        return video.title.toLowerCase().indexOf(term) !== -1;
      });
    }

    // Filter by category
    if (selectedCategory) {
      filtered = filtered.filter(function (video) {
        return video.category === selectedCategory;
      });
    }

    // Filter by department
    if (selectedDepartment) {
      filtered = filtered.filter(function (video) {
        return video.department === selectedDepartment;
      });
    }

    // Filter by duration
    if (durationFilter !== "all") {
      const self = this;
      filtered = filtered.filter(function (video) {
        const durationInSeconds = self._parseDuration(video.duration);

        if (durationFilter === "short") {
          return durationInSeconds < 300; // Less than 5 minutes
        } else if (durationFilter === "medium") {
          return durationInSeconds >= 300 && durationInSeconds < 900; // 5-15 minutes
        } else if (durationFilter === "long") {
          return durationInSeconds >= 900; // 15+ minutes
        }

        return true;
      });
    }

    // Sort videos
    const self = this;
    filtered.sort(function (a, b) {
      let comparison = 0;

      if (sortBy === "title") {
        comparison = a.title.localeCompare(b.title);
      } else if (sortBy === "duration") {
        comparison =
          self._parseDuration(a.duration) - self._parseDuration(b.duration);
      } else if (sortBy === "category") {
        comparison = (a.category || "").localeCompare(b.category || "");
      } else if (sortBy === "department") {
        comparison = (a.department || "").localeCompare(b.department || "");
      }

      return sortDirection === "asc" ? comparison : -comparison;
    });

    this.setState({ filteredVideos: filtered });
  };

  public render(): React.ReactElement<IVideoHubProps> {
    // Extract isRTL from props
    const { isRTL } = this.props;
    const { filteredVideos, sortDirection } = this.state;
    const cardW = 220;

    // Apply RTL classes conditionally
    const containerClass = isRTL
      ? `${styles.videoHub} ${styles.rtl}`
      : styles.videoHub;
    const rowClass = isRTL
      ? `${styles.controlsRow} ${styles.controlsRowRtl}`
      : styles.controlsRow;
    const itemClass = isRTL
      ? `${styles.controlItem} ${styles.controlItemRtl}`
      : styles.controlItem;
    const gridClass = isRTL
      ? `${styles.videoGrid} ${styles.videoGridRtl}`
      : styles.videoGrid;

    const sortOptions: IDropdownOption[] = [
      { key: "title", text: "Title" },
      { key: "duration", text: "Duration" },
      { key: "category", text: "Category" },
      { key: "department", text: "Department" },
    ];

    const durationOptions: IChoiceGroupOption[] = [
      { key: "all", text: "All Durations" },
      { key: "short", text: "Short (<5 min)" },
      { key: "medium", text: "Medium (5-15 min)" },
      { key: "long", text: "Long (>15 min)" },
    ];

    return (
      <div className={containerClass}>
        <div className={styles.controlsContainer}>
          <div className={rowClass}>
            <div className={itemClass}>
              <TextField
                placeholder="Search videos..."
                onChange={this._handleSearch}
                ariaLabel="Search videos"
              />
            </div>
            <div className={itemClass}>
              <Dropdown
                label="Category"
                selectedKey={this.state.selectedCategory}
                options={this.state.categoryOptions}
                onChange={this._handleCategoryChange}
              />
            </div>
            <div className={itemClass}>
              <Dropdown
                label="Department"
                selectedKey={this.state.selectedDepartment}
                options={this.state.departmentOptions}
                onChange={this._handleDepartmentChange}
              />
            </div>
          </div>

          <div className={rowClass}>
            <div aria-label="Filter by duration">
              <ChoiceGroup
                options={durationOptions}
                selectedKey={this.state.durationFilter}
                onChange={this._handleDurationFilterChange}
                label="Duration"
              />
            </div>
          </div>

          <div className={rowClass}>
            <div className={itemClass}>
              <Dropdown
                label="Sort by"
                selectedKey={this.state.sortBy}
                options={sortOptions}
                onChange={this._handleSortChange}
              />
            </div>
            <div className={itemClass}>
              <DefaultButton
                text={`Order: ${sortDirection === "asc" ? "A to Z" : "Z to A"}`}
                onClick={this._toggleSortDirection}
                className={styles.sortButton}
              />
            </div>
          </div>
        </div>

        {filteredVideos.length === 0 ? (
          <div className={styles.noResults}>
            No videos match your search criteria.
          </div>
        ) : (
          <div className={gridClass}>
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
