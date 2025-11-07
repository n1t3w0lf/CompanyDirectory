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
import { IconButton } from '@fluentui/react/lib/Button';
import { Panel } from '@fluentui/react/lib/Panel';
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
    setIsFilterPanelOpen(false);
    await loadInitialUsers();
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

  return (
    <div className={styles.peopleDirectory}>
      <Stack tokens={{ childrenGap: 20 }}>
        {/* Header */}
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Stack.Item grow>
            <Text variant="xxLarge" block>People Directory</Text>
            {totalUsers > 0 && (
              <Stack horizontal tokens={{ childrenGap: 20 }} style={{ marginTop: 8 }}>
                <Stack>
                  <Text variant="small" style={{ color: '#666' }}>
                    {totalUsers.toLocaleString()} people
                  </Text>
                </Stack>
                {departments.length > 0 && (
                  <Stack>
                    <Text variant="small" style={{ color: '#666' }}>
                      {departments.length} departments
                    </Text>
                  </Stack>
                )}
                {locations.length > 0 && (
                  <Stack>
                    <Text variant="small" style={{ color: '#666' }}>
                      {locations.length} office locations
                    </Text>
                  </Stack>
                )}
                {cities.length > 0 && (
                  <Stack>
                    <Text variant="small" style={{ color: '#666' }}>
                      {cities.length} cities
                    </Text>
                  </Stack>
                )}
              </Stack>
            )}
          </Stack.Item>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <IconButton
              iconProps={{ iconName: 'Filter' }}
              title="Advanced filters"
              ariaLabel="Advanced filters"
              onClick={() => setIsFilterPanelOpen(true)}
              text={getActiveFilterCount() > 0 ? `${getActiveFilterCount()} active` : undefined}
            />
            <IconButton
              iconProps={{ iconName: 'Refresh' }}
              title="Reload users"
              ariaLabel="Reload users"
              onClick={() => loadInitialUsers()}
            />
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
          onChange={(_, newValue) => handleSearchChange(newValue)}
          onClear={() => handleSearchChange('')}
          value={searchText}
          disabled={loading}
          className={styles.searchBox}
        />

        {/* Active Filters Display */}
        {getActiveFilterCount() > 0 && (
          <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
            {activeFilters.department && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={() => handleApplyFilters({ ...activeFilters, department: undefined })}
                dismissButtonAriaLabel="Remove filter"
                styles={{ root: { marginBottom: 0 } }}
              >
                Department: {activeFilters.department}
              </MessageBar>
            )}
            {activeFilters.officeLocation && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={() => handleApplyFilters({ ...activeFilters, officeLocation: undefined })}
                dismissButtonAriaLabel="Remove filter"
                styles={{ root: { marginBottom: 0 } }}
              >
                Location: {activeFilters.officeLocation}
              </MessageBar>
            )}
            {activeFilters.city && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={() => handleApplyFilters({ ...activeFilters, city: undefined })}
                dismissButtonAriaLabel="Remove filter"
                styles={{ root: { marginBottom: 0 } }}
              >
                City: {activeFilters.city}
              </MessageBar>
            )}
            {activeFilters.country && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={() => handleApplyFilters({ ...activeFilters, country: undefined })}
                dismissButtonAriaLabel="Remove filter"
                styles={{ root: { marginBottom: 0 } }}
              >
                Country: {activeFilters.country}
              </MessageBar>
            )}
            {activeFilters.jobTitle && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={() => handleApplyFilters({ ...activeFilters, jobTitle: undefined })}
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
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError('')}>
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
                  {users.map(user => (
                    <UserCard
                      key={user.id}
                      user={user}
                      onClick={() => handleUserClick(user)}
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
                      fontSize={props.fontSize}
                      fontColor={props.fontColor}
                      iconSize={props.iconSize}
                      iconColor={props.iconColor}
                    />
                  ))}
                </div>
              </>
            )}
          </>
        )}
      </Stack>

      {/* Advanced Filter Panel */}
      <Panel
        isOpen={isFilterPanelOpen}
        onDismiss={() => setIsFilterPanelOpen(false)}
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
          onDismiss={() => {
            setIsPanelOpen(false);
            setSelectedUser(null);
          }}
        />
      )}
    </div>
  );
};
