import * as React from 'react';
import { useState, useEffect, useCallback, useRef } from 'react';
import { IPeopleDirectoryProps } from './IPeopleDirectoryProps';
import { IUserProfile } from '../../../models/IUserProfile';
import { ISyncStatus } from '../../../services/SyncService';
import { Constants } from '../../../models/Constants';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Panel } from '@fluentui/react/lib/Panel';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { UserCard } from './UserCard';
import { UserDetailsPanel } from './UserDetailsPanel';
import { SyncStatusBanner } from './SyncStatusBanner';
import { AdvancedFilterPanel, IAdvancedFilters } from './AdvancedFilterPanel';
import styles from './PeopleDirectory.module.scss';

export const PeopleDirectory: React.FC<IPeopleDirectoryProps> = (props) => {
  // State management
  const [searchText, setSearchText] = useState<string>('');
  const [users, setUsers] = useState<IUserProfile[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');
  const [selectedUser, setSelectedUser] = useState<IUserProfile | null>(null);
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
  const [isFilterPanelOpen, setIsFilterPanelOpen] = useState<boolean>(false);

  // Filter options
  const [departments, setDepartments] = useState<string[]>([]);
  const [locations, setLocations] = useState<string[]>([]);
  const [cities, setCities] = useState<string[]>([]);
  const [countries, setCountries] = useState<string[]>([]);
  const [jobTitles, setJobTitles] = useState<string[]>([]);

  // Active filters
  const [activeFilters, setActiveFilters] = useState<IAdvancedFilters>({});

  // Letter index
  const [selectedLetter, setSelectedLetter] = useState<string | null>(null);

  // Sync status
  const [syncStatus, setSyncStatus] = useState<ISyncStatus | null>(null);
  const [totalUsers, setTotalUsers] = useState<number>(0);

  const searchTimeoutRef = useRef<number | null>(null);

  // Initialize: Check sync status and load initial users
  useEffect(() => {
    initializeData();
  }, []);

  const initializeData = async (): Promise<void> => {
    try {
      setLoading(true);
      await props.peopleService.initialize();

      // Check if initial sync is needed
      const syncNeeded = await props.syncService.isInitialSyncNeeded();

      if (syncNeeded) {
        // Show sync banner
        setSyncStatus({
          isRunning: false,
          totalUsers: 0,
          processedUsers: 0,
          currentBatch: 0,
          totalBatches: 0,
          errors: [],
          progress: 0
        });
        setLoading(false);
      } else {
        // Load initial 30 users
        await loadInitialUsers();

        // Load filter options in background
        loadFilterOptions();
      }
    } catch (err) {
      console.error('Error initializing:', err);
      setError('Failed to initialize people directory. Please refresh the page.');
      setLoading(false);
    }
  };

  /**
   * Load first 30 users
   */
  const loadInitialUsers = async (): Promise<void> => {
    try {
      setLoading(true);
      const initialUsers = await props.peopleService.getInitialUsers(30);
      setUsers(initialUsers);

      // Get total count
      const count = await props.peopleService.getTotalUserCount();
      setTotalUsers(count);

      setLoading(false);
    } catch (err) {
      console.error('Error loading initial users:', err);
      setError('Failed to load users. Please try again.');
      setLoading(false);
    }
  };

  /**
   * Load filter options
   */
  const loadFilterOptions = async (): Promise<void> => {
    try {
      const [depts, locs, citiesData, countriesData, titlesData] = await Promise.all([
        props.peopleService.getDepartments(),
        props.peopleService.getOfficeLocations(),
        props.peopleService.getCities(),
        props.peopleService.getCountries(),
        props.peopleService.getJobTitles()
      ]);

      setDepartments(depts);
      setLocations(locs);
      setCities(citiesData);
      setCountries(countriesData);
      setJobTitles(titlesData);
    } catch (err) {
      console.error('Error loading filter options:', err);
    }
  };

  /**
   * Handle initial sync start
   */
  const handleStartSync = async (): Promise<void> => {
    try {
      setSyncStatus({
        isRunning: true,
        totalUsers: 0,
        processedUsers: 0,
        currentBatch: 0,
        totalBatches: 0,
        errors: [],
        progress: 0
      });

      await props.syncService.performInitialSync((status) => {
        setSyncStatus(status);
      });

      // Sync complete, load users
      await loadInitialUsers();
      loadFilterOptions();

    } catch (err) {
      console.error('Sync error:', err);
      setError('Sync failed. Please try again.');
      setSyncStatus(props.syncService.getSyncStatus());
    }
  };

  /**
   * Handle sync cancellation
   */
  const handleCancelSync = (): void => {
    props.syncService.cancelSync();
    setSyncStatus(props.syncService.getSyncStatus());
  };

  /**
   * Debounced search
   */
  const handleSearchChange = useCallback((newValue?: string): void => {
    const value = newValue || '';
    setSearchText(value);

    if (searchTimeoutRef.current) {
      clearTimeout(searchTimeoutRef.current);
    }

    if (value.length < Constants.MIN_SEARCH_LENGTH) {
      // Reload initial users when search is cleared
      loadInitialUsers();
      return;
    }

    searchTimeoutRef.current = setTimeout(() => {
      performSearch(value);
    }, Constants.SEARCH_DEBOUNCE_MS) as unknown as number;
  }, []);

  /**
   * Perform search
   */
  const performSearch = async (searchValue: string): Promise<void> => {
    if (searchValue.length < Constants.MIN_SEARCH_LENGTH) {
      return;
    }

    setLoading(true);
    setError('');

    try {
      const filters = {
        searchText: searchValue,
        ...activeFilters
      };

      const result = await props.peopleService.getFilteredUsers(filters, 100);
      setUsers(result);
    } catch (err) {
      console.error('Search error:', err);
      setError('Failed to search users. Please try again.');
      setUsers([]);
    } finally {
      setLoading(false);
    }
  };

  /**
   * Apply advanced filters
   */
  const handleApplyFilters = async (filters: IAdvancedFilters): Promise<void> => {
    setActiveFilters(filters);
    setIsFilterPanelOpen(false);
    setLoading(true);

    try {
      const combinedFilters = {
        searchText: searchText.length >= Constants.MIN_SEARCH_LENGTH ? searchText : undefined,
        ...filters
      };

      const result = await props.peopleService.getFilteredUsers(combinedFilters, 100);
      setUsers(result);
    } catch (err) {
      console.error('Filter error:', err);
      setError('Failed to apply filters. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  /**
   * Clear all filters
   */
  const handleClearFilters = async (): Promise<void> => {
    setActiveFilters({});
    setSearchText('');
    setSelectedLetter(null);
    setIsFilterPanelOpen(false);
    await loadInitialUsers();
  };

  /**
   * Handle letter index click
   */
  const handleLetterClick = async (letter: string): Promise<void> => {
    // Toggle letter selection
    const newLetter = selectedLetter === letter ? null : letter;
    setSelectedLetter(newLetter);

    if (!newLetter) {
      // If letter is deselected, reload initial users
      await loadInitialUsers();
      return;
    }

    setLoading(true);
    try {
      // Filter users by letter
      const allUsers = await props.peopleService.getFilteredUsers(activeFilters, 1000);
      const filteredUsers = allUsers.filter(user =>
        user.displayName.toUpperCase().startsWith(newLetter)
      );
      setUsers(filteredUsers);
    } catch (err) {
      console.error('Letter filter error:', err);
      setError('Failed to filter by letter. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  /**
   * Handle user card click
   */
  const handleUserClick = async (user: IUserProfile): Promise<void> => {
    setSelectedUser(user);
    setIsPanelOpen(true);

    // Verify user data in background with Entra ID
    try {
      const freshUser = await props.peopleService.getUserById(user.id, true);
      if (freshUser) {
        setSelectedUser(freshUser);
        // Update the user in the list
        setUsers(prevUsers =>
          prevUsers.map(u => u.id === freshUser.id ? freshUser : u)
        );
      }
    } catch (err) {
      console.error('Error refreshing user:', err);
    }
  };

  /**
   * Count active filters
   */
  const getActiveFilterCount = (): number => {
    return Object.values(activeFilters).filter(v => v).length;
  };

  // Callback handlers for dropdowns
  const handleDepartmentFilterChange = useCallback((_: React.FormEvent<HTMLDivElement>, option?: { key: string; text: string }) => {
    const newFilters = { ...activeFilters, department: option?.key || undefined };
    handleApplyFilters(newFilters);
  }, [activeFilters]);

  const handleLocationFilterChange = useCallback((_: React.FormEvent<HTMLDivElement>, option?: { key: string; text: string }) => {
    const newFilters = { ...activeFilters, officeLocation: option?.key || undefined };
    handleApplyFilters(newFilters);
  }, [activeFilters]);

  const handleCityFilterChange = useCallback((_: React.FormEvent<HTMLDivElement>, option?: { key: string; text: string }) => {
    const newFilters = { ...activeFilters, city: option?.key || undefined };
    handleApplyFilters(newFilters);
  }, [activeFilters]);

  // Callback for search box
  const handleSearchBoxChange = useCallback((_: React.ChangeEvent<HTMLInputElement> | undefined, newValue?: string) => {
    handleSearchChange(newValue);
  }, []);

  const handleSearchBoxClear = useCallback(() => {
    handleSearchChange('');
  }, []);

  // Callback for dismissing panel
  const handlePanelDismiss = useCallback(() => {
    setIsPanelOpen(false);
    setSelectedUser(null);
  }, []);

  // Callbacks for filter dismissals
  const handleDismissDepartmentFilter = useCallback(() => {
    handleApplyFilters({ ...activeFilters, department: undefined });
  }, [activeFilters]);

  const handleDismissLocationFilter = useCallback(() => {
    handleApplyFilters({ ...activeFilters, officeLocation: undefined });
  }, [activeFilters]);

  const handleDismissCityFilter = useCallback(() => {
    handleApplyFilters({ ...activeFilters, city: undefined });
  }, [activeFilters]);

  const handleDismissCountryFilter = useCallback(() => {
    handleApplyFilters({ ...activeFilters, country: undefined });
  }, [activeFilters]);

  const handleDismissJobTitleFilter = useCallback(() => {
    handleApplyFilters({ ...activeFilters, jobTitle: undefined });
  }, [activeFilters]);

  const handleDismissError = useCallback(() => {
    setError('');
  }, []);

  const handleDismissFilterPanel = useCallback(() => {
    setIsFilterPanelOpen(false);
  }, []);

  // Build webpart background style
  const webpartBackgroundStyle: React.CSSProperties = {};
  if (props.webpartBackgroundImage) {
    webpartBackgroundStyle.backgroundImage = `url(${props.webpartBackgroundImage})`;
    webpartBackgroundStyle.backgroundSize = 'cover';
    webpartBackgroundStyle.backgroundPosition = 'center';
    webpartBackgroundStyle.backgroundRepeat = 'no-repeat';
  } else if (props.webpartBackgroundColor) {
    webpartBackgroundStyle.backgroundColor = props.webpartBackgroundColor;
  }

  return (
    <div className={styles.peopleDirectory} style={webpartBackgroundStyle}>
      <Stack tokens={{ childrenGap: 20 }}>
        {/* Header */}
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
            <Stack.Item grow>
              <Text
                block
                style={{
                  fontSize: `${props.headingFontSize}px`,
                  color: props.headingFontColor,
                  fontWeight: 600,
                  marginBottom: props.subtextText ? 8 : 0
                }}
              >
                {props.headingText}
              </Text>
              {props.subtextText && (
                <Text
                  block
                  style={{
                    fontSize: `${props.subtextFontSize}px`,
                    color: props.subtextFontColor
                  }}
                >
                  {props.subtextText}
                </Text>
              )}
            </Stack.Item>
          </Stack>

          {/* Filter Dropdowns */}
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            <Stack.Item styles={{ root: { minWidth: 200 } }}>
              <Dropdown
                placeholder="Filter by Department"
                options={[
                  { key: '', text: 'All Departments' },
                  ...departments.map(d => ({ key: d, text: d }))
                ]}
                selectedKey={activeFilters.department || ''}
                onChange={handleDepartmentFilterChange}
                styles={{ dropdown: { width: 200 } }}
              />
            </Stack.Item>
            <Stack.Item styles={{ root: { minWidth: 200 } }}>
              <Dropdown
                placeholder="Filter by Office Location"
                options={[
                  { key: '', text: 'All Locations' },
                  ...locations.map(l => ({ key: l, text: l }))
                ]}
                selectedKey={activeFilters.officeLocation || ''}
                onChange={handleLocationFilterChange}
                styles={{ dropdown: { width: 200 } }}
              />
            </Stack.Item>
            <Stack.Item styles={{ root: { minWidth: 200 } }}>
              <Dropdown
                placeholder="Filter by City"
                options={[
                  { key: '', text: 'All Cities' },
                  ...cities.map(c => ({ key: c, text: c }))
                ]}
                selectedKey={activeFilters.city || ''}
                onChange={handleCityFilterChange}
                styles={{ dropdown: { width: 200 } }}
              />
            </Stack.Item>
            {totalUsers > 0 && (
              <Stack.Item grow>
                <Text variant="small" style={{ color: '#666', paddingTop: 8 }}>
                  {totalUsers.toLocaleString()} people
                </Text>
              </Stack.Item>
            )}
          </Stack>
        </Stack>

        {/* Sync Status Banner */}
        {syncStatus && (
          <SyncStatusBanner
            syncStatus={syncStatus}
            onStartSync={handleStartSync}
            onCancelSync={handleCancelSync}
          />
        )}

        {/* Search Box */}
        <SearchBox
          placeholder={Constants.MSG_SEARCH_PLACEHOLDER}
          onChange={handleSearchBoxChange}
          onClear={handleSearchBoxClear}
          value={searchText}
          disabled={loading}
          className={styles.searchBox}
        />

        {/* Letter Index */}
        {props.showLetterIndex && (
          <div className={styles.letterIndex}>
            <Stack horizontal tokens={{ childrenGap: 4 }} wrap>
              {Array.from('ABCDEFGHIJKLMNOPQRSTUVWXYZ').map((letter) => {
                const handleClick = (): void => {
                  handleLetterClick(letter);
                };
                const handleKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    handleLetterClick(letter);
                  }
                };
                return (
                  <div
                    key={letter}
                    className={`${styles.letterButton} ${selectedLetter === letter ? styles.letterButtonActive : ''}`}
                    onClick={handleClick}
                    role="button"
                    tabIndex={0}
                    onKeyDown={handleKeyDown}
                  >
                    {letter}
                  </div>
                );
              })}
            </Stack>
          </div>
        )}

        {/* Active Filters Display */}
        {getActiveFilterCount() > 0 && (
          <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
            {activeFilters.department && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={handleDismissDepartmentFilter}
                dismissButtonAriaLabel="Remove filter"
                styles={{ root: { marginBottom: 0 } }}
              >
                Department: {activeFilters.department}
              </MessageBar>
            )}
            {activeFilters.officeLocation && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={handleDismissLocationFilter}
                dismissButtonAriaLabel="Remove filter"
                styles={{ root: { marginBottom: 0 } }}
              >
                Location: {activeFilters.officeLocation}
              </MessageBar>
            )}
            {activeFilters.city && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={handleDismissCityFilter}
                dismissButtonAriaLabel="Remove filter"
                styles={{ root: { marginBottom: 0 } }}
              >
                City: {activeFilters.city}
              </MessageBar>
            )}
            {activeFilters.country && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={handleDismissCountryFilter}
                dismissButtonAriaLabel="Remove filter"
                styles={{ root: { marginBottom: 0 } }}
              >
                Country: {activeFilters.country}
              </MessageBar>
            )}
            {activeFilters.jobTitle && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={handleDismissJobTitleFilter}
                dismissButtonAriaLabel="Remove filter"
                styles={{ root: { marginBottom: 0 } }}
              >
                Job Title: {activeFilters.jobTitle}
              </MessageBar>
            )}
          </Stack>
        )}

        {/* Error Message */}
        {error && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={handleDismissError}>
            {error}
          </MessageBar>
        )}

        {/* Loading Spinner */}
        {loading && (
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
            <Spinner size={SpinnerSize.large} label={Constants.MSG_LOADING} />
          </Stack>
        )}

        {/* Results */}
        {!loading && (
          <>
            {users.length === 0 ? (
              <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
                <Text variant="large">{Constants.MSG_NO_RESULTS}</Text>
                {searchText || getActiveFilterCount() > 0 ? (
                  <Text variant="medium" style={{ marginTop: 8, color: '#666' }}>
                    Try adjusting your search or filters
                  </Text>
                ) : (
                  <Text variant="medium" style={{ marginTop: 8, color: '#666' }}>
                    Start typing to search for people
                  </Text>
                )}
              </Stack>
            ) : (
              <>
                <Text variant="medium">
                  {searchText || getActiveFilterCount() > 0 ? 'Found' : 'Showing'} {users.length} {users.length === 1 ? 'person' : 'people'}
                </Text>
                <div className={styles.userGrid}>
                  {users.map(user => {
                    const handleClick = (): void => {
                      handleUserClick(user);
                    };
                    return (
                      <UserCard
                        key={user.id}
                        user={user}
                        onClick={handleClick}
                        showEmail={props.showEmail}
                      showJobTitle={props.showJobTitle}
                      showDepartment={props.showDepartment}
                      showOfficeLocation={props.showOfficeLocation}
                      showBusinessPhones={props.showBusinessPhones}
                      showMobilePhone={props.showMobilePhone}
                      showCity={props.showCity}
                      showCountry={props.showCountry}
                      showCompanyName={props.showCompanyName}
                      showEmployeeId={props.showEmployeeId}
                      profileNameFontSize={props.profileNameFontSize}
                      profileNameFontColor={props.profileNameFontColor}
                      jobTitleFontSize={props.jobTitleFontSize}
                      jobTitleFontColor={props.jobTitleFontColor}
                      jobTitleBold={props.jobTitleBold}
                      profilePropertiesFontSize={props.profilePropertiesFontSize}
                      profilePropertiesFontColor={props.profilePropertiesFontColor}
                      propertyDisplayOrder={props.propertyDisplayOrder}
                      textEllipsisLength={props.textEllipsisLength}
                      iconSize={props.iconSize}
                      iconColor={props.iconColor}
                      profileCardBackgroundColor={props.profileCardBackgroundColor}
                      profileCardBackgroundImage={props.profileCardBackgroundImage}
                      showProfilePicture={props.showProfilePicture}
                      />
                    );
                  })}
                </div>
              </>
            )}
          </>
        )}
      </Stack>

      {/* Advanced Filter Panel */}
      <Panel
        isOpen={isFilterPanelOpen}
        onDismiss={handleDismissFilterPanel}
        headerText="Advanced Filters"
        closeButtonAriaLabel="Close"
        isLightDismiss
      >
        <AdvancedFilterPanel
          departments={departments}
          locations={locations}
          cities={cities}
          countries={countries}
          jobTitles={jobTitles}
          currentFilters={activeFilters}
          onApplyFilters={handleApplyFilters}
          onClearFilters={handleClearFilters}
        />
      </Panel>

      {/* User Details Panel */}
      {selectedUser && (
        <UserDetailsPanel
          user={selectedUser}
          isOpen={isPanelOpen}
          onDismiss={handlePanelDismiss}
        />
      )}
    </div>
  );
};
