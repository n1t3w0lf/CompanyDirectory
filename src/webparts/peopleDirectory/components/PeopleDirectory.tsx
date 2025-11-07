import * as React from 'react';
import { useState, useEffect, useCallback, useRef } from 'react';
import { IPeopleDirectoryProps } from './IPeopleDirectoryProps';
import { IUserProfile, ISearchFilter } from '../../../models/IUserProfile';
import { Constants } from '../../../models/Constants';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { IconButton } from '@fluentui/react/lib/Button';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Panel } from '@fluentui/react/lib/Panel';
import { UserCard } from './UserCard';
import { UserDetailsPanel } from './UserDetailsPanel';
import styles from './PeopleDirectory.module.scss';

export const PeopleDirectory: React.FC<IPeopleDirectoryProps> = (props) => {
  const [searchText, setSearchText] = useState<string>('');
  const [users, setUsers] = useState<IUserProfile[]>([]);
  const [filteredUsers, setFilteredUsers] = useState<IUserProfile[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [selectedUser, setSelectedUser] = useState<IUserProfile | null>(null);
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
  const [isFilterPanelOpen, setIsFilterPanelOpen] = useState<boolean>(false);
  const [departments, setDepartments] = useState<string[]>([]);
  const [locations, setLocations] = useState<string[]>([]);
  const [selectedDepartment, setSelectedDepartment] = useState<string>('');
  const [selectedLocation, setSelectedLocation] = useState<string>('');
  const [hasSearched, setHasSearched] = useState<boolean>(false);

  const searchTimeoutRef = useRef<number | null>(null);

  // Initialize service and load filters
  useEffect(() => {
    initializeData();
  }, []);

  const initializeData = async (): Promise<void> => {
    try {
      await props.peopleService.initialize();

      // Load filter options in background
      const [depts, locs] = await Promise.all([
        props.peopleService.getDepartments(),
        props.peopleService.getOfficeLocations()
      ]);

      setDepartments(depts);
      setLocations(locs);
    } catch (err) {
      console.error('Error initializing:', err);
      setError('Failed to initialize people directory. Please refresh the page.');
    }
  };

  // Debounced search
  const handleSearchChange = useCallback((newValue?: string): void => {
    const value = newValue || '';
    setSearchText(value);

    if (searchTimeoutRef.current) {
      clearTimeout(searchTimeoutRef.current);
    }

    if (value.length < Constants.MIN_SEARCH_LENGTH) {
      setUsers([]);
      setFilteredUsers([]);
      setHasSearched(false);
      return;
    }

    searchTimeoutRef.current = setTimeout(() => {
      performSearch(value);
    }, Constants.SEARCH_DEBOUNCE_MS);
  }, []);

  // Perform search
  const performSearch = async (searchValue: string): Promise<void> => {
    if (searchValue.length < Constants.MIN_SEARCH_LENGTH) {
      return;
    }

    setLoading(true);
    setError('');
    setHasSearched(true);

    try {
      const filter: ISearchFilter = {
        searchText: searchValue,
        department: selectedDepartment || undefined,
        officeLocation: selectedLocation || undefined
      };

      const result = await props.peopleService.searchUsers(filter);
      setUsers(result.users);
      applyLocalFilters(result.users);
    } catch (err) {
      console.error('Search error:', err);
      setError('Failed to search users. Please try again.');
      setUsers([]);
      setFilteredUsers([]);
    } finally {
      setLoading(false);
    }
  };

  // Apply local filters (department, location)
  const applyLocalFilters = (userList: IUserProfile[]): void => {
    let filtered = [...userList];

    if (selectedDepartment) {
      filtered = filtered.filter(u => u.department === selectedDepartment);
    }

    if (selectedLocation) {
      filtered = filtered.filter(u => u.officeLocation === selectedLocation);
    }

    setFilteredUsers(filtered);
  };

  // Handle filter changes
  useEffect(() => {
    applyLocalFilters(users);
  }, [selectedDepartment, selectedLocation, users]);

  // Handle user card click
  const handleUserClick = async (user: IUserProfile): Promise<void> => {
    setSelectedUser(user);
    setIsPanelOpen(true);

    // Verify user data in background
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

  const departmentOptions: IDropdownOption[] = [
    { key: '', text: 'All Departments' },
    ...departments.map(d => ({ key: d, text: d }))
  ];

  const locationOptions: IDropdownOption[] = [
    { key: '', text: 'All Locations' },
    ...locations.map(l => ({ key: l, text: l }))
  ];

  return (
    <div className={styles.peopleDirectory}>
      <Stack tokens={{ childrenGap: 20 }}>
        {/* Header */}
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Stack.Item grow>
            <Text variant="xxLarge" block>{props.title}</Text>
            {props.description && (
              <Text variant="medium" block style={{ color: '#666' }}>
                {props.description}
              </Text>
            )}
          </Stack.Item>
          <IconButton
            iconProps={{ iconName: 'Filter' }}
            title="Filter options"
            ariaLabel="Filter options"
            onClick={() => setIsFilterPanelOpen(true)}
          />
        </Stack>

        {/* Search Box */}
        <SearchBox
          placeholder={Constants.MSG_SEARCH_PLACEHOLDER}
          onChange={(_, newValue) => handleSearchChange(newValue)}
          onClear={() => handleSearchChange('')}
          value={searchText}
          disabled={loading}
          className={styles.searchBox}
        />

        {/* Active Filters */}
        {(selectedDepartment || selectedLocation) && (
          <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
            {selectedDepartment && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={() => setSelectedDepartment('')}
                dismissButtonAriaLabel="Remove filter"
                styles={{ root: { marginBottom: 0 } }}
              >
                Department: {selectedDepartment}
              </MessageBar>
            )}
            {selectedLocation && (
              <MessageBar
                messageBarType={MessageBarType.info}
                onDismiss={() => setSelectedLocation('')}
                dismissButtonAriaLabel="Remove filter"
                styles={{ root: { marginBottom: 0 } }}
              >
                Location: {selectedLocation}
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
        {!loading && hasSearched && (
          <>
            {filteredUsers.length === 0 ? (
              <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
                <Text variant="large">{Constants.MSG_NO_RESULTS}</Text>
              </Stack>
            ) : (
              <>
                <Text variant="medium">
                  Found {filteredUsers.length} {filteredUsers.length === 1 ? 'person' : 'people'}
                </Text>
                <div className={styles.userGrid}>
                  {filteredUsers.map(user => (
                    <UserCard
                      key={user.id}
                      user={user}
                      onClick={() => handleUserClick(user)}
                    />
                  ))}
                </div>
              </>
            )}
          </>
        )}

        {/* Instructions when no search */}
        {!loading && !hasSearched && (
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
            <Text variant="large" style={{ color: '#666' }}>
              Start typing to search for people in your organization
            </Text>
          </Stack>
        )}
      </Stack>

      {/* Filter Panel */}
      <Panel
        isOpen={isFilterPanelOpen}
        onDismiss={() => setIsFilterPanelOpen(false)}
        headerText="Filter Options"
        closeButtonAriaLabel="Close"
        isLightDismiss
      >
        <Stack tokens={{ childrenGap: 20 }} styles={{ root: { marginTop: 20 } }}>
          <Dropdown
            label="Department"
            options={departmentOptions}
            selectedKey={selectedDepartment}
            onChange={(_, option) => setSelectedDepartment(option?.key as string || '')}
          />
          <Dropdown
            label="Office Location"
            options={locationOptions}
            selectedKey={selectedLocation}
            onChange={(_, option) => setSelectedLocation(option?.key as string || '')}
          />
        </Stack>
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
