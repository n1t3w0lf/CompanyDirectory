/**
 * Application constants
 */

export class Constants {
  // SharePoint List Configuration
  public static readonly LIST_TITLE = 'PeopleDirectoryCache';
  public static readonly LIST_MAX_ITEMS = 5000;

  // Cache Configuration
  public static readonly INDEXEDDB_NAME = 'PeopleDirectoryDB';
  public static readonly INDEXEDDB_VERSION = 1;
  public static readonly INDEXEDDB_STORE = 'UserProfiles';
  public static readonly CLIENT_CACHE_TTL = 30 * 60 * 1000; // 30 minutes
  public static readonly LIST_CACHE_TTL = 24 * 60 * 60 * 1000; // 24 hours

  // Search Configuration
  public static readonly SEARCH_DEBOUNCE_MS = 3000; // 3 seconds
  public static readonly MIN_SEARCH_LENGTH = 2;
  public static readonly DEFAULT_PAGE_SIZE = 50;
  public static readonly VIRTUAL_SCROLL_ITEM_HEIGHT = 120;

  // Graph API Configuration
  public static readonly GRAPH_BATCH_SIZE = 20;
  public static readonly GRAPH_SELECT_FIELDS = [
    'id',
    'userPrincipalName',
    'displayName',
    'givenName',
    'surname',
    'mail',
    'jobTitle',
    'department',
    'officeLocation',
    'businessPhones',
    'mobilePhone',
    'city',
    'country',
    'companyName',
    'preferredLanguage',
    'employeeId'
  ].join(',');

  // Photo Configuration
  public static readonly PHOTO_SIZE = '96x96';
  public static readonly DEFAULT_PHOTO_PLACEHOLDER = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iOTYiIGhlaWdodD0iOTYiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PHJlY3Qgd2lkdGg9Ijk2IiBoZWlnaHQ9Ijk2IiBmaWxsPSIjZjNmMmYxIi8+PHRleHQgeD0iNTAlIiB5PSI1MCUiIGZvbnQtc2l6ZT0iNDAiIGZpbGw9IiM2MDVlNWMiIGRvbWluYW50LWJhc2VsaW5lPSJtaWRkbGUiIHRleHQtYW5jaG9yPSJtaWRkbGUiPj88L3RleHQ+PC9zdmc+';

  // Error Messages
  public static readonly ERROR_GRAPH_API = 'Unable to fetch user data from Microsoft Graph. Please try again.';
  public static readonly ERROR_LIST_PROVISION = 'Failed to create or access the cache list. Please contact your administrator.';
  public static readonly ERROR_NETWORK = 'Network error occurred. Please check your connection and try again.';
  public static readonly ERROR_PERMISSION = 'Insufficient permissions. Please contact your administrator.';
  public static readonly ERROR_GENERIC = 'An unexpected error occurred. Please try again.';

  // UI Messages
  public static readonly MSG_NO_RESULTS = 'No people found matching your search criteria.';
  public static readonly MSG_LOADING = 'Loading people directory...';
  public static readonly MSG_SEARCHING_AD = 'Searching Active Directory...';
  public static readonly MSG_SEARCH_PLACEHOLDER = 'Search by name, email, or department...';
  public static readonly MSG_FILTER_DEPT = 'Filter by department';
  public static readonly MSG_FILTER_LOCATION = 'Filter by location';
}
