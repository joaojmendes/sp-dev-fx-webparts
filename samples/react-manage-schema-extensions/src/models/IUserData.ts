export interface IUserProfile {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  businessPhones?: string[];
  mobilePhone?: string;
  companyName?: string;
  employeeId?: string;
  managerId?: string;
  aboutMe?: string;
  skills?: string[];
  photoUrl?: string;
  userType?: string;
}

export interface IDirectReport {
  id: string;
  displayName: string;
  jobTitle?: string;
  department?: string;
  mail?: string;
  userPrincipalName: string;
  photoUrl?: string;
  location?: string;
  skills?: string[];
  aboutMe?: string;
  phone?: string;
  userType?: string;
}

export interface IManager {
  id: string;
  displayName: string;
  jobTitle?: string;
  department?: string;
  mail?: string;
  userPrincipalName: string;
  photoUrl?: string;
  location?: string;
  skills?: string[];
  aboutMe?: string;
  phone?: string;
  userType?: string;
  totalDirectReports?: number;
}

export interface IUserData {
  user: IUserProfile;
  manager?: IManager;
  managers: IManager[]; // Array of all managers in the hierarchy
  totalDirectReports: number
  peers: IDirectReport[];
}

export interface IOrganizationNode {
  id: string;
  displayName: string;
  jobTitle?: string;
  department?: string;
  mail?: string;
  userPrincipalName: string;
  photoUrl?: string;
  location?: string;
  skills?: string[];
  aboutMe?: string;
  phone?: string;
  managerId?: string; // direct manager ID
  managers: IManager[]; // Array of all managers in the hierarchy
  level: number;
  totalDirectReports?: number;
  isExpanded?: boolean;
  hasDirectReports?: boolean;
  userType?: string;
}
