import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { IPeopleDirectoryProps } from './IPeopleDirectoryProps';
import { IUserProfile } from '../../../models/IUserProfile';
import { ISyncStatus } from '../../../services/SyncService';
import { Constants } from '../../../models/Constants';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { UserCard } from './UserCard';
import { UserDetailsPanel } from './UserDetailsPanel';
import { SyncStatusBanner } from './SyncStatusBanner';
import { LetterIndex } from './LetterIndex';
import { Pagination } from './Pagination';
import { IAdvancedFilters } from './AdvancedFilterPanel';
import styles from './PeopleDirectory.module.scss';

export const PeopleDirectory: React.FC<IPeopleDirectoryProps> = (props) => {
  // State management
  const [searchText, setSearchText] = useState<string>('');
  const [users, setUsers] = useState<IUserProfile[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');
  const [selectedUser, setSelectedUser] = useState<IUserProfile | null>(null);
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);

  // Filter options
  const [departments, setDepartments] = useState<string[]>([]);
  const [locations, setLocations] = useState<string[]>([]);
  const [cities, setCities] = useState<string[]>([]);

  // Active filters
  const [activeFilters, setActiveFilters] = useState<IAdvancedFilters>({});

  // Letter index filter
  const [selectedLetter, setSelectedLetter] = useState<string | null>(null);

  // Sync status
  const [syncStatus, setSyncStatus] = useState<ISyncStatus | null>(null);

  // Client-side pagination
  const [currentPage, setCurrentPage] = useState<number>(1);

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
        // Load filter options in background (needed for dropdowns regardless)
        loadFilterOptions();

        if (props.showPeopleOnStart) {
          // Load the configured number of people
          await loadInitialUsers();
        } else {
          // Startup list disabled: show the search prompt instead
          setUsers([]);
          setLoading(false);
        }
      }
    } catch (err) {
      console.error('Error initializing:', err);
      setError('Failed to initialize people directory. Please refresh the page.');
      setLoading(false);
    }
  };

  /**
   * Load the configured number of people (props.initialPeopleCount) for the startup view
   */
  const loadInitialUsers = useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      setCurrentPage(1);
      const initialUsers = await props.peopleService.getInitialUsers(props.initialPeopleCount);
      setUsers(initialUsers);

      setLoading(false);
    } catch (err) {
      console.error('Error loading initial users:', err);
      setError('Failed to load users. Please try again.');
      setLoading(false);
    }
  }, [props.peopleService, props.initialPeopleCount]);

  /**
   * Load filter options
   */
  const loadFilterOptions = useCallback(async (): Promise<void> => {
    try {
      const [depts, locs, citiesData] = await Promise.all([
        props.peopleService.getDepartments(),
        props.peopleService.getOfficeLocations(),
        props.peopleService.getCities()
      ]);

      setDepartments(depts);
      setLocations(locs);
      setCities(citiesData);
    } catch (err) {
      console.error('Error loading filter options:', err);
    }
  }, [props.peopleService]);

  /**
   * Handle initial sync start
   */
  const handleStartSync = useCallback(async (): Promise<void> => {
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

      // Sync complete: load filter options, then restore the startup view
      loadFilterOptions();
      if (props.showPeopleOnStart) {
        await loadInitialUsers();
      } else {
        setUsers([]);
      }

    } catch (err) {
      console.error('Sync error:', err);
      setError('Sync failed. Please try again.');
      setSyncStatus(props.syncService.getSyncStatus());
    }
  }, [props.syncService, props.showPeopleOnStart, loadInitialUsers, loadFilterOptions]);

  /**
   * Handle sync cancellation
   */
  const handleCancelSync = useCallback((): void => {
    props.syncService.cancelSync();
    setSyncStatus(props.syncService.getSyncStatus());
  }, [props.syncService]);

  /**
   * Manual search: Check list first, then Entra ID
   */
  const handleManualSearch = useCallback(async (): Promise<void> => {
    if (!searchText || searchText.trim().length < Constants.MIN_SEARCH_LENGTH) {
      setError('Please enter at least 3 characters to search');
      return;
    }

    setLoading(true);
    setError('');

    try {
      const result = await props.peopleService.manualSearch(searchText.trim());

      setUsers(result.users);
      setCurrentPage(1);
      setSelectedLetter(null);

      if (!result.success) {
        setError(result.message);
      }
    } catch (err) {
      console.error('Search error:', err);
      setError('An error occurred while searching. Please try again.');
      setUsers([]);
    } finally {
      setLoading(false);
    }
  }, [searchText, props.peopleService]);

  /**
   * Handle search text change (no auto-search)
   */
  const handleSearchChange = useCallback((newValue?: string): void => {
    const value = newValue || '';
    setSearchText(value);
    setError('');

    // If search is cleared, restore the startup view
    if (value.length === 0) {
      setCurrentPage(1);
      setSelectedLetter(null);
      if (props.showPeopleOnStart) {
        loadInitialUsers();
      } else {
        setUsers([]);
      }
    }
  }, [loadInitialUsers, props.showPeopleOnStart]);

  /**
   * Apply advanced filters
   */
  const handleApplyFilters = useCallback(async (filters: IAdvancedFilters): Promise<void> => {
    setActiveFilters(filters);
    setLoading(true);

    try {
      const combinedFilters = {
        searchText: searchText.length >= Constants.MIN_SEARCH_LENGTH ? searchText : undefined,
        ...filters
      };

      const result = await props.peopleService.getFilteredUsers(combinedFilters, 100);
      setUsers(result);
      setCurrentPage(1);
      setSelectedLetter(null);
    } catch (err) {
      console.error('Filter error:', err);
      setError('Failed to apply filters. Please try again.');
    } finally {
      setLoading(false);
    }
  }, [searchText, props.peopleService]);

  /**
   * Clear all filters
   */
  const handleClearFilters = useCallback(async (): Promise<void> => {
    setActiveFilters({});
    setSearchText('');
    setSelectedLetter(null);
    setCurrentPage(1);
    if (props.showPeopleOnStart) {
      await loadInitialUsers();
    } else {
      setUsers([]);
    }
  }, [loadInitialUsers, props.showPeopleOnStart]);

  /**
   * Handle user card click
   */
  const handleUserClick = useCallback(async (user: IUserProfile): Promise<void> => {
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
  }, [props.peopleService]);

  /**
   * Count active filters
   */
  const getActiveFilterCount = (): number => {
    return Object.values(activeFilters).filter(v => v).length;
  };

  /**
   * Get available letters from current users
   */
  const getAvailableLetters = (): Set<string> => {
    const letters = new Set<string>();
    users.forEach(user => {
      const firstLetter = user.displayName.charAt(0).toUpperCase();
      if (/[A-Z]/.test(firstLetter)) {
        letters.add(firstLetter);
      }
    });
    return letters;
  };

  /**
   * Filter users by selected letter
   */
  const getFilteredUsersByLetter = (): IUserProfile[] => {
    if (!selectedLetter) {
      return users;
    }
    return users.filter(user =>
      user.displayName.charAt(0).toUpperCase() === selectedLetter
    );
  };

  /**
   * Handle letter selection
   */
  const handleLetterSelect = useCallback((letter: string | null): void => {
    setSelectedLetter(letter);
    setCurrentPage(1);
  }, []);

  /**
   * Handle pagination page change
   */
  const handlePageChange = useCallback((page: number): void => {
    setCurrentPage(page);
  }, []);

  /**
   * Inline filter change handlers
   */
  const handleDepartmentChange = useCallback((_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    const newFilters = { ...activeFilters, department: option?.key ? String(option.key) : undefined };
    handleApplyFilters(newFilters);
  }, [activeFilters, handleApplyFilters]);

  const handleOfficeLocationChange = useCallback((_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    const newFilters = { ...activeFilters, officeLocation: option?.key ? String(option.key) : undefined };
    handleApplyFilters(newFilters);
  }, [activeFilters, handleApplyFilters]);

  const handleCityChange = useCallback((_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    const newFilters = { ...activeFilters, city: option?.key ? String(option.key) : undefined };
    handleApplyFilters(newFilters);
  }, [activeFilters, handleApplyFilters]);

  const handleClearAllFilters = useCallback(async (): Promise<void> => {
    await handleClearFilters();
  }, [handleClearFilters]);

  /**
   * SearchBox handlers
   */
  const handleSearchBoxChange = useCallback((_: React.ChangeEvent<HTMLInputElement> | undefined, newValue?: string): void => {
    handleSearchChange(newValue);
  }, [handleSearchChange]);

  const handleSearchBoxClear = useCallback((): void => {
    handleSearchChange('');
  }, [handleSearchChange]);

  /**
   * Panel dismiss handler
   */
  const handleDetailsPanelDismiss = useCallback((): void => {
    setIsPanelOpen(false);
    setSelectedUser(null);
  }, []);

  /**
   * Error dismiss handler
   */
  const handleErrorDismiss = useCallback((): void => {
    setError('');
  }, []);

  // Build webpart container style
  const webpartContainerStyle: React.CSSProperties = React.useMemo(() => {
    const style: React.CSSProperties = {
      ...(props.webpartBackgroundColor && { backgroundColor: props.webpartBackgroundColor }),
      ...(props.webpartBackgroundImage && {
        backgroundImage: `url("${props.webpartBackgroundImage}")`,
        backgroundSize: 'cover',
        backgroundPosition: 'center',
        backgroundRepeat: 'no-repeat'
      })
    };

    // Debug logging
    if (props.webpartBackgroundImage || props.webpartBackgroundColor) {
      console.log('Webpart background config:', {
        backgroundColor: props.webpartBackgroundColor,
        backgroundImage: props.webpartBackgroundImage,
        styleApplied: style
      });
    }

    return style;
  }, [props.webpartBackgroundColor, props.webpartBackgroundImage]);

  return (
    <div className={styles.peopleDirectory} style={webpartContainerStyle}>
      <Stack tokens={{ childrenGap: 20 }}>
        {/* Header */}
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Stack.Item grow>
            <Text
              variant="xxLarge"
              block
              style={{
                fontSize: props.headingFontSize,
                color: props.headingFontColor
              }}
            >
              {props.headingText}
            </Text>
            {props.subtextText && (
              <Text
                variant="medium"
                block
                style={{
                  fontSize: props.subtextFontSize,
                  color: props.subtextFontColor
                }}
              >
                {props.subtextText}
              </Text>
            )}
          </Stack.Item>
        </Stack>

        {/* Sync Status Banner */}
        {syncStatus && (
          <SyncStatusBanner
            syncStatus={syncStatus}
            onStartSync={handleStartSync}
            onCancelSync={handleCancelSync}
          />
        )}

        {/* Search Box with Search Button */}
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
          <Stack.Item grow>
            <SearchBox
              placeholder={Constants.MSG_SEARCH_PLACEHOLDER}
              onChange={handleSearchBoxChange}
              onClear={handleSearchBoxClear}
              onSearch={handleManualSearch}
              value={searchText}
              disabled={loading}
              className={styles.searchBox}
            />
          </Stack.Item>
          <PrimaryButton
            text={props.searchButtonText || 'Search'}
            onClick={handleManualSearch}
            disabled={loading || !searchText || searchText.trim().length < Constants.MIN_SEARCH_LENGTH}
            iconProps={{ iconName: 'Search' }}
            styles={{
              root: {
                backgroundColor: props.searchButtonColor || undefined,
                borderColor: props.searchButtonColor || undefined
              },
              rootHovered: {
                backgroundColor: props.searchButtonHoverColor || undefined,
                borderColor: props.searchButtonHoverColor || undefined
              },
              label: {
                fontSize: props.searchButtonTextSize ? `${props.searchButtonTextSize}px` : undefined
              }
            }}
          />
        </Stack>

        {/* Inline Filters */}
        <Stack tokens={{ childrenGap: 12 }}>
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            <Stack.Item styles={{ root: { width: '200px' } }}>
              <Dropdown
                placeholder="All Departments"
                label="Department"
                options={[
                  { key: '', text: 'All Departments' },
                  ...departments.map(d => ({ key: d, text: d }))
                ]}
                selectedKey={activeFilters.department || ''}
                onChange={handleDepartmentChange}
                disabled={loading}
              />
            </Stack.Item>
            <Stack.Item styles={{ root: { width: '200px' } }}>
              <Dropdown
                placeholder="All Locations"
                label="Office Location"
                options={[
                  { key: '', text: 'All Locations' },
                  ...locations.map(l => ({ key: l, text: l }))
                ]}
                selectedKey={activeFilters.officeLocation || ''}
                onChange={handleOfficeLocationChange}
                disabled={loading}
              />
            </Stack.Item>
            <Stack.Item styles={{ root: { width: '200px' } }}>
              <Dropdown
                placeholder="All Cities"
                label="City"
                options={[
                  { key: '', text: 'All Cities' },
                  ...cities.map(c => ({ key: c, text: c }))
                ]}
                selectedKey={activeFilters.city || ''}
                onChange={handleCityChange}
                disabled={loading}
              />
            </Stack.Item>
          </Stack>
          {getActiveFilterCount() > 0 && (
            <PrimaryButton
              text="Clear All Filters"
              onClick={handleClearAllFilters}
              iconProps={{ iconName: 'ClearFilter' }}
              styles={{ root: { width: 'fit-content' } }}
            />
          )}
        </Stack>

        {/* Letter Index */}
        {props.showLetterIndex && !loading && users.length > 0 && (
          <LetterIndex
            selectedLetter={selectedLetter}
            onLetterSelect={handleLetterSelect}
            availableLetters={getAvailableLetters()}
          />
        )}

        {/* Error Message */}
        {error && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={handleErrorDismiss}>
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
              (() => {
                const hasQuery = !!searchText || getActiveFilterCount() > 0;
                return (
                  <div className={styles.emptyState}>
                    <Icon
                      iconName={hasQuery ? 'SearchIssue' : 'People'}
                      className={styles.emptyStateIcon}
                    />
                    <Text block className={styles.emptyStateText}>
                      {hasQuery ? Constants.MSG_NO_RESULTS : 'Search for people'}
                    </Text>
                    <Text block className={styles.emptyStateSubtext}>
                      {hasQuery
                        ? 'Try adjusting your search or filters'
                        : 'Type a name, email, or department, then press Search.'}
                    </Text>
                  </div>
                );
              })()
            ) : (
              <>
                {(() => {
                  const filteredUsers = getFilteredUsersByLetter();
                  const pageSize = props.paginationSize;
                  // Clamp against the current result size so a runtime page-size
                  // change (property pane) can never strand us on an empty page.
                  const pageCount = Math.max(1, Math.ceil(filteredUsers.length / pageSize));
                  const safePage = Math.min(currentPage, pageCount);
                  const pagedUsers = filteredUsers.slice(
                    (safePage - 1) * pageSize,
                    safePage * pageSize
                  );
                  return (
                    <>
                      <Text variant="medium">
                        {`${searchText || getActiveFilterCount() > 0 || selectedLetter ? 'Found' : 'Showing'} ${filteredUsers.length} ${filteredUsers.length === 1 ? 'person' : 'people'}${selectedLetter ? ` starting with "${selectedLetter}"` : ''}`}
                      </Text>
                      {filteredUsers.length === 0 ? (
                        <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
                          <Text variant="medium" style={{ color: '#666' }}>
                            {`No people found starting with "${selectedLetter}"`}
                          </Text>
                        </Stack>
                      ) : (
                        <>
                        <div className={styles.userGrid}>
                          {pagedUsers.map(user => (
                    <UserCard
                      key={user.id}
                      user={user}
                      onUserClick={handleUserClick}
                      profileNameFontSize={props.profileNameFontSize}
                      profileNameFontColor={props.profileNameFontColor}
                      jobTitleFontSize={props.jobTitleFontSize}
                      jobTitleFontColor={props.jobTitleFontColor}
                      jobTitleBold={props.jobTitleBold}
                      profilePropertiesFontSize={props.profilePropertiesFontSize}
                      profilePropertiesFontColor={props.profilePropertiesFontColor}
                      iconSize={props.iconSize}
                      iconColor={props.iconColor}
                      profileCardBackgroundColor={props.profileCardBackgroundColor}
                      profileCardBackgroundImage={props.profileCardBackgroundImage}
                      showProfilePicture={props.showProfilePicture}
                      textEllipsisLength={props.textEllipsisLength}
                      showJobTitle={props.showJobTitle}
                      showDepartment={props.showDepartment}
                      showOfficeLocation={props.showOfficeLocation}
                      showEmail={props.showEmail}
                      showBusinessPhones={props.showBusinessPhones}
                      showMobilePhone={props.showMobilePhone}
                      showCity={props.showCity}
                      showCountry={props.showCountry}
                      showCompanyName={props.showCompanyName}
                      showEmployeeId={props.showEmployeeId}
                    />
                  ))}
                        </div>
                        <Pagination
                          currentPage={safePage}
                          totalItems={filteredUsers.length}
                          pageSize={pageSize}
                          onPageChange={handlePageChange}
                        />
                        </>
                      )}
                    </>
                  );
                })()}
              </>
            )}
          </>
        )}
      </Stack>

      {/* User Details Panel */}
      {selectedUser && (
        <UserDetailsPanel
          user={selectedUser}
          isOpen={isPanelOpen}
          onDismiss={handleDetailsPanelDismiss}
          profileNameFontSize={props.profileNameFontSize}
          profileNameFontColor={props.profileNameFontColor}
          jobTitleFontSize={props.jobTitleFontSize}
          jobTitleFontColor={props.jobTitleFontColor}
          jobTitleBold={props.jobTitleBold}
          profilePropertiesFontSize={props.profilePropertiesFontSize}
          profilePropertiesFontColor={props.profilePropertiesFontColor}
          iconSize={props.iconSize}
          iconColor={props.iconColor}
          showProfilePicture={props.showProfilePicture}
          showEmail={true}
          showJobTitle={true}
          showDepartment={true}
          showOfficeLocation={true}
          showBusinessPhones={true}
          showMobilePhone={true}
          showCity={true}
          showCountry={true}
          showCompanyName={true}
          showEmployeeId={true}
        />
      )}
    </div>
  );
};
