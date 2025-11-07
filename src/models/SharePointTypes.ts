/**
 * SharePoint list item types
 */

export interface ISharePointListItem {
  Id: number;
  Title: string;
  [key: string]: unknown;
}

export interface ISharePointField {
  InternalName: string;
  [key: string]: unknown;
}

export interface ISharePointList {
  fields: {
    select: (fields: string) => () => Promise<ISharePointField[]>;
    addText: (name: string, options: unknown) => Promise<unknown>;
    addDateTime: (name: string, options: unknown) => Promise<unknown>;
    addNumber: (name: string, options: unknown) => Promise<unknown>;
    getByInternalNameOrTitle: (name: string) => { update: (options: unknown) => Promise<unknown> };
  };
  [key: string]: unknown;
}

export interface IGraphUser {
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
  [key: string]: unknown;
}

export interface IGraphResponse {
  value: IGraphUser[];
  '@odata.count'?: number;
  '@odata.nextLink'?: string;
}

export interface IBatchResponse {
  responses: Array<{
    id: string;
    status: number;
    body: IGraphUser;
  }>;
}

export interface IUserCacheListItem extends ISharePointListItem {
  UserId: string;
  UserPrincipalName: string;
  GivenName?: string;
  Surname?: string;
  Email?: string;
  JobTitle?: string;
  Department?: string;
  OfficeLocation?: string;
  BusinessPhones?: string;
  MobilePhone?: string;
  City?: string;
  Country?: string;
  CompanyName?: string;
  PhotoUrl?: string;
  LastVerified?: string;
  AccessCount?: number;
}
