// src/webparts/videoFilter/components/VideoFilter.tsx
import * as React from "react";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { DefaultButton } from "office-ui-fabric-react/lib/Button";
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from "office-ui-fabric-react/lib/ChoiceGroup";
import {
  VideoHubStateService,
  IVideoFilterOptions,
} from "../../../shared/VideoHubStateService";
import styles from "../../videoHub/components/VideoHub.module.scss";

export interface IVideoFilterProps {
  isRTL?: boolean;
}

export interface IVideoFilterState {
  filterOptions: IVideoFilterOptions;
  categoryOptions: IDropdownOption[];
  departmentOptions: IDropdownOption[];
}

export default class VideoFilter extends React.Component<
  IVideoFilterProps,
  IVideoFilterState
> {
  private _stateService: VideoHubStateService;

  constructor(props: IVideoFilterProps) {
    super(props);

    this._stateService = VideoHubStateService.getInstance();

    // Extract unique categories and departments
    const categoryOptions = this._getUniqueOptions("category");
    const departmentOptions = this._getUniqueOptions("department");

    this.state = {
      filterOptions: this._stateService.getFilterOptions(),
      categoryOptions,
      departmentOptions,
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
    // Update local state when shared state changes
    this.setState({
      filterOptions: this._stateService.getFilterOptions(),
      categoryOptions: this._getUniqueOptions("category"),
      departmentOptions: this._getUniqueOptions("department"),
    });
  }

  private _getUniqueOptions(field: string): IDropdownOption[] {
    const videos = this._stateService.getVideos();
    const uniqueValues = {};

    for (let i = 0; i < videos.length; i++) {
      const value = videos[i][field];
      if (value) {
        uniqueValues[value] = true;
      }
    }

    const options: IDropdownOption[] = [];
    for (const key in uniqueValues) {
      if (uniqueValues.hasOwnProperty(key)) {
        options.push({
          key: key,
          text: key,
        });
      }
    }

    // Add "All" option at the beginning
    const allOption: IDropdownOption = { key: "", text: "All" };
    const result: IDropdownOption[] = [allOption];

    for (let i = 0; i < options.length; i++) {
      result.push(options[i]);
    }

    return result;
  }

  private _handleSearch = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this._stateService.updateFilterOptions({ searchTerm: newValue || "" });
  };

  private _handleCategoryChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this._stateService.updateFilterOptions({
        selectedCategory: option.key as string,
      });
    }
  };

  private _handleDepartmentChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this._stateService.updateFilterOptions({
        selectedDepartment: option.key as string,
      });
    }
  };

  private _handleDurationFilterChange = (
    ev?: React.FormEvent<HTMLElement | HTMLInputElement>,
    option?: IChoiceGroupOption
  ): void => {
    if (option) {
      this._stateService.updateFilterOptions({ durationFilter: option.key });
    }
  };

  private _handleSortChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this._stateService.updateFilterOptions({ sortBy: option.key as string });
    }
  };

  private _toggleSortDirection = (): void => {
    const newDirection =
      this.state.filterOptions.sortDirection === "asc" ? "desc" : "asc";
    this._stateService.updateFilterOptions({ sortDirection: newDirection });
  };

  private _resetFilters = (): void => {
    this._stateService.updateFilterOptions({
      searchTerm: "",
      selectedCategory: "",
      selectedDepartment: "",
      durationFilter: "all",
      sortBy: "title",
      sortDirection: "asc",
    });
  };

  public render(): React.ReactElement<IVideoFilterProps> {
    const { isRTL } = this.props;
    const { filterOptions } = this.state;

    // Apply RTL classes conditionally
    const rowClass = isRTL
      ? `${styles.controlsRow} ${styles.controlsRowRtl}`
      : styles.controlsRow;
    const itemClass = isRTL
      ? `${styles.controlItem} ${styles.controlItemRtl}`
      : styles.controlItem;

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
      <div className={styles.filterContainer}>
        <div className={styles.filterHeader}>
          <h3 className={styles.filterTitle}>Video Filters</h3>
          <DefaultButton
            text="Reset Filters"
            onClick={this._resetFilters}
            className={styles.resetButton}
          />
        </div>

        <div className={styles.controlsContainer}>
          <div className={rowClass}>
            <div className={itemClass}>
              <TextField
                placeholder="Search videos..."
                onChange={this._handleSearch}
                value={filterOptions.searchTerm}
                ariaLabel="Search videos"
              />
            </div>
            <div className={itemClass}>
              <Dropdown
                label="Category"
                selectedKey={filterOptions.selectedCategory}
                options={this.state.categoryOptions}
                onChange={this._handleCategoryChange}
              />
            </div>
            <div className={itemClass}>
              <Dropdown
                label="Department"
                selectedKey={filterOptions.selectedDepartment}
                options={this.state.departmentOptions}
                onChange={this._handleDepartmentChange}
              />
            </div>
          </div>

          <div className={rowClass}>
            <div aria-label="Filter by duration">
              <ChoiceGroup
                options={durationOptions}
                selectedKey={filterOptions.durationFilter}
                onChange={this._handleDurationFilterChange}
                label="Duration"
              />
            </div>
          </div>

          <div className={rowClass}>
            <div className={itemClass}>
              <Dropdown
                label="Sort by"
                selectedKey={filterOptions.sortBy}
                options={sortOptions}
                onChange={this._handleSortChange}
              />
            </div>
            <div className={itemClass}>
              <DefaultButton
                text={`Order: ${
                  filterOptions.sortDirection === "asc" ? "A to Z" : "Z to A"
                }`}
                onClick={this._toggleSortDirection}
                className={styles.sortButton}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
