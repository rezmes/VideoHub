// src/webparts/videoFilter/components/VideoFilter.tsx
import * as React from "react";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { DefaultButton } from "office-ui-fabric-react/lib/Button";
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from "office-ui-fabric-react/lib/ChoiceGroup";
import {
  VideoHubStateService,
  IVideoFilterOptions,
} from "../../../shared/VideoHubStateService";
import GenericDropdown from "../../../shared/components/GenericDropdown";
import styles from './VideoFilter.module.scss';
import * as mystrings from "VideoFilterWebPartStrings";

// Temporary strings object until localization is properly set up
const strings = {
  FilterTitle: mystrings.FilterTitle,
  ResetFilters: mystrings.ResetFilters,
  SearchPlaceholder: mystrings.SearchPlaceholder,
  CategoryLabel: mystrings.CategoryLabel,
  DepartmentLabel: mystrings.DepartmentLabel,
  DurationLabel: mystrings.DurationLabel,
  AllDurations: mystrings.AllDurations,
  ShortVideos: mystrings.ShortVideos,
  MediumVideos: mystrings.MediumVideos,
  LongVideos: mystrings.LongVideos,
  SortByLabel: mystrings.SortByLabel,
  SortDirectionAsc: mystrings.SortDirectionAsc,
  SortDirectionDesc: mystrings.SortDirectionDesc
};

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
    console.log('Search term:', newValue);
    this._stateService.updateFilterOptions({ searchTerm: newValue || "" });
  };

  private _handleCategoryChange = (option?: IDropdownOption): void => {
    console.log('Category selected:', option);
    if (option) {
      this._stateService.updateFilterOptions({
        selectedCategory: option.key as string,
      });
      console.log('After update - filterOptions:', this._stateService.getFilterOptions());
    }
  };

  private _handleDepartmentChange = (option?: IDropdownOption): void => {
    console.log('Department selected:', option);
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
    console.log('Duration selected:', option);
    if (option) {
      this._stateService.updateFilterOptions({ durationFilter: option.key });
    }
  };

  private _handleSortChange = (option?: IDropdownOption): void => {
    console.log('Sort option selected:', option);
    if (option) {
      this._stateService.updateFilterOptions({ sortBy: option.key as string });
    }
  };

  private _toggleSortDirection = (): void => {
    const newDirection =
      this.state.filterOptions.sortDirection === "asc" ? "desc" : "asc";
    console.log('Toggling sort direction to:', newDirection);
    this._stateService.updateFilterOptions({ sortDirection: newDirection });
  };

  private _resetFilters = (): void => {
    console.log('Resetting filters');
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

    console.log('Rendering with filterOptions:', filterOptions);
    console.log('Category options:', this.state.categoryOptions);
    console.log('Department options:', this.state.departmentOptions);

    // Apply RTL classes conditionally using string indexing to avoid TypeScript errors
    const rowClass = `${styles.controlsRow}${isRTL ? ' ' + styles['controlsRowRtl'] : ''}`;
    const itemClass = `${styles.controlItem}${isRTL ? ' ' + styles['controlItemRtl'] : ''}`;

    const sortOptions: IDropdownOption[] = [
      { key: "title", text: "Title" },
      { key: "duration", text: "Duration" },
      { key: "category", text: "Category" },
      { key: "department", text: "Department" },
    ];

    const durationOptions: IChoiceGroupOption[] = [
      { key: "all", text: strings.AllDurations },
      { key: "short", text: strings.ShortVideos },
      { key: "medium", text: strings.MediumVideos },
      { key: "long", text: strings.LongVideos },
    ];

    return (
      <div className={styles.filterContainer}>
        <div className={styles.filterHeader}>
          <h3 className={styles.filterTitle}>{strings.FilterTitle}</h3>
          <DefaultButton
            text={strings.ResetFilters}
            onClick={this._resetFilters}
            className={styles.resetButton}
          />
        </div>

        <div className={styles.controlsContainer}>
          <div className={rowClass}>
            <div className={`${itemClass} ${styles.searchBox}`}>
              <TextField
                placeholder={strings.SearchPlaceholder}
                onChange={this._handleSearch}
                value={filterOptions.searchTerm}
                ariaLabel="Search videos"
              />
            </div>
            <div className={itemClass}>
              <GenericDropdown
                label={strings.CategoryLabel}
                selectedKey={filterOptions.selectedCategory}
                options={this.state.categoryOptions}
                onChanged={this._handleCategoryChange}
                className={styles.dropdown}
              />
            </div>
            <div className={itemClass}>
              <GenericDropdown
                label={strings.DepartmentLabel}
                selectedKey={filterOptions.selectedDepartment}
                options={this.state.departmentOptions}
                onChanged={this._handleDepartmentChange}
                className={styles.dropdown}
              />
            </div>
          </div>

          <div className={rowClass}>
            <div aria-label="Filter by duration">
              <ChoiceGroup
                options={durationOptions}
                selectedKey={filterOptions.durationFilter}
                onChange={this._handleDurationFilterChange}
                label={strings.DurationLabel}
              />
            </div>
          </div>

          <div className={rowClass}>
            <div className={itemClass}>
              <GenericDropdown
                label={strings.SortByLabel}
                selectedKey={filterOptions.sortBy}
                options={sortOptions}
                onChanged={this._handleSortChange}
                className={styles.dropdown}
              />
            </div>
            <div className={itemClass}>
              <DefaultButton
                text={filterOptions.sortDirection === "asc" 
                  ? strings.SortDirectionAsc 
                  : strings.SortDirectionDesc}
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
