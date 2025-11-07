/**
 * User profile interface matching Microsoft Graph User entity
 * https://docs.microsoft.com/en-us/graph/api/resources/user
 */
export interface IUserProfile {
  id: string;
  userPrincipalName: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  businessPhones?: string[];
  mobilePhone?: string;
  city?: string;
  country?: string;
  companyName?: string;
  preferredLanguage?: string;
  employeeId?: string;
  manager?: IUserProfile;
  photoUrl?: string;
  lastVerified?: Date;
  accessCount?: number;
}

/**
 * Search filter criteria
 */
export interface ISearchFilter {
  searchText: string;
  department?: string;
  officeLocation?: string;
  jobTitle?: string;
}

/**
 * Cache entry metadata
 */
export interface ICacheEntry<T> {
  data: T;
  timestamp: number;
  ttl: number; // Time to live in milliseconds
}

/**
 * Pagination parameters
 */
export interface IPaginationParams {
  pageSize: number;
  pageNumber: number;
  totalItems?: number;
}

/**
 * Search result with pagination info
 */
export interface ISearchResult {
  users: IUserProfile[];
  totalCount: number;
  hasMore: boolean;
  nextLink?: string;
}

/**
 * Department summary for filter dropdown
 */
export interface IDepartment {
  name: string;
  count: number;
}

/**
 * Error details
 */
export interface IErrorDetails {
  message: string;
  code?: string;
  timestamp: Date;
  context?: string;
}
