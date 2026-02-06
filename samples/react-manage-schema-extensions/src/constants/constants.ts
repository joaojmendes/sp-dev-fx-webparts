
export const AZURE_API_ENDPOINT = "https://aiagent.spteckapps.com/api/aiassistant";
export const PAGE_SIZE = 100;

// Microsoft Graph API supported sortable properties for User objects
// Based on: https://docs.microsoft.com/en-us/graph/query-parameters#orderby-parameter
// Note: Properties marked as "Advanced" require ConsistencyLevel: eventual header and $count=true
export const SORTABLE_USER_PROPERTIES: string[] = [
  'displayName',
  'mail', 
  'userPrincipalName',
  'givenName',
  'surname',
  'userType',
  'department'  // Requires advanced query parameters
];