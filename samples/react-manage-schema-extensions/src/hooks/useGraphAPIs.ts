import * as React from "react";

import { ErrorType, useLogging } from "@spteck/m365-hooks";
import { IDirectReport, IManager, IUserProfile } from "../models/IUserData";

import { BaseComponentContext } from "@microsoft/sp-component-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { SORTABLE_USER_PROPERTIES } from "../constants/constants";
import { User } from "@microsoft/microsoft-graph-types";
import { useUtils } from "../utils/useUtils";


// Types and interfaces
export interface IUseGraphAPIsProps {
  context: BaseComponentContext;
}

export interface IUsersPageResultSimplified {
  users: IUserProfile[];
  hasMore: boolean;
  nextPageToken?: string;
  totalCount?: number;
}

export interface IUseGraphAPIsResult {
  getUserProfile: (userId?: string) => Promise<IUserProfile>;
  getUserManager: (userId?: string) => Promise<IManager | undefined>;
  getAllManagers: (userId?: string) => Promise<IManager[]>;
  getUserDirectReports: (
    userId?: string,
    pageSize?: number,
    nextPageToken?: string
  ) => Promise<{
    directReports: IDirectReport[];
    hasMore: boolean;
    nextPageToken?: string;
    totalCount?: number;
  }>;
  getUserPeers: (
    userId?: string,
    pageSize?: number,
    nextPageToken?: string
  ) => Promise<{
    peers: IDirectReport[];
    hasMore: boolean;
    nextPageToken?: string;
    totalCount?: number;
  }>;
  getUserPhoto: (userId?: string) => Promise<string | undefined>;
  searchUsers: (searchTerm: string, top?: number) => Promise<IDirectReport[]>;
  hasDirectReports: (userId: string) => Promise<boolean>;
  getTotalDirectReports: (userId?: string) => Promise<number>;
  getUsers: (
    pageToken?: string,
    pageSize?: number,
    filter?: string,
    sortColumn?: string,
    sortDirection?: "ascending" | "descending"
  ) => Promise<IUsersPageResultSimplified>;
}

// Base fields to select for user queries
const USER_BASE_FIELDS =
  "id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,businessPhones,mobilePhone,companyName,employeeId,userType";
const USER_EXTENDED_FIELDS = `${USER_BASE_FIELDS},aboutMe,skills`;

export const useGraphAPIs = ({
  context,
}: IUseGraphAPIsProps): IUseGraphAPIsResult => {
  const { logError } = useLogging();
  const { blobToBase64, sanitizeUserData } = useUtils();
  // Create Graph client function
  const getGraphClient =
    React.useCallback(async (): Promise<MSGraphClientV3> => {
      if (!context) {
        throw new Error("Context not available");
      }
      return await context.msGraphClientFactory.getClient("3");
    }, [context]);

  /**
   * Helper function to handle Graph API errors consistently
   */
  const handleGraphError = React.useCallback(
    (
      error: unknown,
      operation: string,
      additionalContext?: Record<string, unknown>
    ): Error => {
      const graphError = error as Error & { code?: string; status?: number };

      console.error(`Graph API Error in ${operation}:`, error);

      logError(
        "useGraphAPIs",
        `Error in ${operation}`,
        graphError,
        ErrorType.SYSTEM,
        sanitizeUserData({
          operation,
          errorCode: graphError.code,
          errorStatus: graphError.status,
          errorMessage: graphError.message,
          ...additionalContext,
        })
      );

      return new Error(`Failed to ${operation}: ${graphError.message}`);
    },
    []
  );

  /**
   * Get user photo as base64 data URL
   */
  const getUserPhoto = React.useCallback(
    async (userId?: string): Promise<string | undefined> => {
      try {
        const endpoint = userId
          ? `/users/${userId}/photo/$value`
          : "/me/photo/$value";
        const client = await getGraphClient();

        const photoBlob = await client.api(endpoint).get();

        if (photoBlob instanceof Blob) {
          return await blobToBase64(photoBlob);
        }

        return undefined;
      } catch (error: unknown) {
        const graphError = error as Error & { code?: string; status?: number };

        // Photo not found is expected and not an error
        if (graphError.code === "ImageNotFound" || graphError.status === 404) {
          return undefined;
        }

        console.warn(
          `Could not fetch photo for user ${userId || "current"}:`,
          error
        );
        return undefined;
      }
    },
    []
  );

  /**
   * Get user profile information
   */
  const getUserProfile = React.useCallback(
    async (userId?: string): Promise<IUserProfile> => {
      try {
        const endpoint = userId ? `/users/${userId}` : "/me";
        const client = await getGraphClient();

        const response: User = await client
          .api(endpoint)
          .select(USER_EXTENDED_FIELDS)
          .get();

        const userProfile: IUserProfile = {
          id: response.id || "",
          displayName: response.displayName || "",
          mail: response.mail || response.userPrincipalName || "",
          userPrincipalName: response.userPrincipalName || "",
          jobTitle: response.jobTitle || "",
          department: response.department || "",
          officeLocation: response.officeLocation || "",
          businessPhones: response.businessPhones || [],
          mobilePhone: response.mobilePhone || "",
          companyName: response.companyName || "",
          employeeId: response.employeeId || "",
          aboutMe: response.aboutMe || "",
          skills: response.skills || [],
          userType: response.userType || "",
        };

        // Fetch photo separately (non-blocking)
        try {
          userProfile.photoUrl = `${
            context.pageContext.web.absoluteUrl
          }//_layouts/15/userphoto.aspx?size=L&username=${
            response.userPrincipalName || ""
          }`;
        } catch {
          // Photo fetch failure is not critical
        }

        return userProfile;
      } catch (error) {
        throw handleGraphError(error, "getUserProfile", {
          userId: userId || "current",
        });
      }
    },
    []
  );

  /**
   * Get user's manager
   */
  const getUserManager = React.useCallback(
    async (userId?: string): Promise<IManager | undefined> => {
      try {
        const endpoint = userId ? `/users/${userId}/manager` : "/me/manager";
        const client = await getGraphClient();

        const response = await client
          .api(endpoint)
          .select(USER_BASE_FIELDS)
          .get();

        if (!response?.id) {
          return undefined;
        }

        const manager: IManager = {
          id: response.id,
          displayName: response.displayName,
          jobTitle: response.jobTitle,
          department: response.department,
          mail: response.mail || response.userPrincipalName || "",
          userPrincipalName: response.userPrincipalName,
        };

        // Fetch photo separately (non-blocking)
        try {
          manager.photoUrl = `${
            context.pageContext.web.absoluteUrl
          }//_layouts/15/userphoto.aspx?size=L&username=${
            response.userPrincipalName || ""
          }`;
        } catch {
          // Photo fetch failure is not critical
        }

        return manager;
      } catch (error: unknown) {
        const graphError = error as Error & { code?: string; status?: number };

        // No manager is expected and not an error
        if (
          graphError.code === "Request_ResourceNotFound" ||
          graphError.status === 404
        ) {
          return undefined;
        }

        throw handleGraphError(error, "getUserManager", {
          userId: userId || "current",
        });
      }
    },
    []
  );

  /**
   * Get the total count of direct reports for a user
   */
  const getTotalDirectReports = React.useCallback(
    async (userId?: string): Promise<number> => {
      try {
        const endpoint = userId
          ? `/users/${userId}/directReports/$count`
          : "/me/directReports/$count";
        const client = await getGraphClient();

        const count = await client
          .api(endpoint)
          .header("ConsistencyLevel", "eventual")
          .get();

        return Number(count) || 0;
      } catch (error: unknown) {
        const graphError = error as Error & { code?: string; status?: number };

        // User not found or no direct reports is expected and not an error
        if (
          graphError.code === "Request_ResourceNotFound" ||
          graphError.status === 404
        ) {
          return 0;
        }

        logError(
          "useGraphAPIs",
          `Could not fetch direct reports count for user ${
            userId || "current"
          }`,
          error as Error,
          ErrorType.SYSTEM,
          sanitizeUserData({
            userId: userId || "current",
            operation: "getTotalDirectReports",
          })
        );
        return 0;
      }
    },
    [getGraphClient]
  );

  /**
   * Get all managers in the hierarchy
   */
  const getAllManagers = React.useCallback(
    async (userId?: string): Promise<IManager[]> => {
      try {
        const endpoint = userId ? `/users/${userId}` : "/me";
        const client = await getGraphClient();

        const response = await client
          .api(endpoint)
          .expand(`manager($levels=max;$select=${USER_BASE_FIELDS})`)
          .select("id,displayName")
          .header("ConsistencyLevel", "eventual")
          .get();

        const managers: IManager[] = [];
        let currentManager = response.manager;

        while (currentManager) {
          const manager: IManager = {
            id: currentManager.id,
            displayName: currentManager.displayName,
            jobTitle: currentManager.jobTitle,
            department: currentManager.department,
            mail: currentManager.mail,
            userPrincipalName: currentManager.userPrincipalName,
          };

          // Fetch photo separately (non-blocking)
          try {
            manager.photoUrl = `${
              context.pageContext.web.absoluteUrl
            }//_layouts/15/userphoto.aspx?size=L&username=${
              currentManager.userPrincipalName || ""
            }`;
          } catch {
            // Photo fetch failure is not critical
          }

          managers.push(manager);
          currentManager = currentManager.manager;
        }

        // Fetch direct reports count for all managers in parallel using the existing function
        if (managers.length > 0) {
          try {
            const countPromises = managers.map(async (manager) => {
              try {
                const count = await getTotalDirectReports(manager.id);
                return { managerId: manager.id, count };
              } catch (error) {
                logError(
                  "useGraphAPIs",
                  `Failed to get direct reports count for manager ${manager.id}`,
                  error as Error,
                  ErrorType.SYSTEM,
                  sanitizeUserData({
                    managerId: manager.id,
                    operation: "getTotalDirectReports",
                  })
                );
                return { managerId: manager.id, count: 0 };
              }
            });

            const countResults = await Promise.all(countPromises);

            // Update managers with their direct reports count
            countResults.forEach(({ managerId, count }) => {
              const manager = managers.find((m) => m.id === managerId);
              if (manager) {
                manager.totalDirectReports = count;
              }
            });
          } catch (error) {
            logError(
              "useGraphAPIs",
              "Error fetching direct reports count for managers",
              error as Error,
              ErrorType.SYSTEM,
              sanitizeUserData({
                operation: "getAllManagers_fetchCounts",
                managersCount: managers.length,
              })
            );
            // Set default count to 0 for all managers if parallel fetch fails
            managers.forEach((manager) => {
              manager.totalDirectReports = 0;
            });
          }
        }

        return managers.reverse();
      } catch (error: unknown) {
        const graphError = error as Error & { code?: string; status?: number };

        // No managers is expected and not an error
        if (
          graphError.code === "Request_ResourceNotFound" ||
          graphError.status === 404
        ) {
          return [];
        }

        throw handleGraphError(error, "getAllManagers", {
          userId: userId || "current",
        });
      }
    },
    [context, getGraphClient, getTotalDirectReports]
  );

  /**
   * Get user's direct reports with paging support
   */
  const getUserDirectReports = React.useCallback(
    async (
      userId?: string,
      pageSize: number = 20,
      nextPageToken?: string
    ): Promise<{
      directReports: IDirectReport[];
      hasMore: boolean;
      nextPageToken?: string;
      totalCount?: number;
    }> => {
      try {
        const client = await getGraphClient();
        let response: {
          value?: Array<{
            id: string;
            displayName: string;
            jobTitle?: string;
            department?: string;
            mail?: string;
            userPrincipalName: string;
          }>;
          "@odata.nextLink"?: string;
          "@odata.count"?: number;
        };

        if (nextPageToken) {
          // Use next page token
          response = await client.api(nextPageToken).get();
        } else {
          // Initial request
          const endpoint = userId
            ? `/users/${userId}/directReports`
            : "/me/directReports";
          response = await client
            .api(endpoint)
            .select(USER_BASE_FIELDS)
            .top(pageSize)
            .count(true)
            .get();
        }

        if (!response?.value?.length) {
          return {
            directReports: [],
            hasMore: false,
            nextPageToken: undefined,
            totalCount: response?.["@odata.count"],
          };
        }

        // Process direct reports with photos
        const directReports = await Promise.allSettled(
          response.value.map(
            async (report: {
              id: string;
              displayName: string;
              jobTitle?: string;
              department?: string;
              mail?: string;
              userPrincipalName: string;
            }): Promise<IDirectReport> => {
              const directReport: IDirectReport = {
                id: report.id,
                displayName: report.displayName,
                jobTitle: report.jobTitle,
                department: report.department,
                mail: report.mail || report.userPrincipalName || "",
                userPrincipalName: report.userPrincipalName,
              };

              // Fetch photo separately (non-blocking)
              try {
                directReport.photoUrl = `${
                  context.pageContext.web.absoluteUrl
                }//_layouts/15/userphoto.aspx?size=L&username=${
                  report.userPrincipalName || ""
                }`;
              } catch {
                // Photo fetch failure is not critical
              }

              return directReport;
            }
          )
        );

        return {
          directReports: directReports
            .filter(
              (result: PromiseSettledResult<IDirectReport>) =>
                result.status === "fulfilled"
            )
            .map(
              (result: PromiseSettledResult<IDirectReport>) =>
                (result as PromiseFulfilledResult<IDirectReport>).value
            ),
          hasMore: Boolean(response?.["@odata.nextLink"]),
          nextPageToken: response?.["@odata.nextLink"],
          totalCount: response?.["@odata.count"],
        };
      } catch (error) {
        throw handleGraphError(error, "getUserDirectReports", {
          userId: userId || "current",
        });
      }
    },
    [context]
  );

  /**
   * Get user's peers (colleagues with the same manager) with pagination support
   */
  const getUserPeers = React.useCallback(
    async (
      userId?: string,
      pageSize: number = 20,
      nextPageToken?: string
    ): Promise<{
      peers: IDirectReport[];
      hasMore: boolean;
      nextPageToken?: string;
      totalCount?: number;
    }> => {
      try {
        const manager = await getUserManager(userId);
        if (!manager) {
          return {
            peers: [],
            hasMore: false,
            nextPageToken: undefined,
            totalCount: 0,
          };
        }

        const result = await getUserDirectReports(
          manager.id,
          pageSize,
          nextPageToken
        );
        const currentUserId = userId || (await getUserProfile()).id;

        // Filter out the current user from the direct reports to get peers
        const peers = result.directReports.filter(
          (report: IDirectReport) => report.id !== currentUserId
        );

        return {
          peers,
          hasMore: result.hasMore,
          nextPageToken: result.nextPageToken,
          totalCount: result.totalCount ? result.totalCount - 1 : undefined, // Subtract 1 for the current user
        };
      } catch (error) {
        logError(
          "useGraphAPIs",
          "Error fetching user peers",
          error as Error,
          ErrorType.SYSTEM,
          sanitizeUserData({
            userId: userId || "current",
            operation: "getUserPeers",
          })
        );
        return {
          peers: [],
          hasMore: false,
          nextPageToken: undefined,
          totalCount: 0,
        };
      }
    },
    [getUserManager, getUserDirectReports, getUserProfile]
  );

  /**
   * Search users in the organization
   */
  const searchUsers = React.useCallback(
    async (searchTerm: string, top: number = 20): Promise<IDirectReport[]> => {
      try {
        const client = await getGraphClient();

        let apiCall = client.api("/users");

        // Build search filter based on term type
        if (
          searchTerm.includes(" eq ") ||
          searchTerm.includes(" and ") ||
          searchTerm.includes(" or ") ||
          searchTerm.includes("contains(") ||
          searchTerm.includes("startswith(")
        ) {
          // OData filter expression - encode special characters
          const encodedSearchTerm = encodeURIComponent(searchTerm);
          apiCall = apiCall.filter(encodedSearchTerm);
        } else {
          // Simple search term - encode special characters and create comprehensive filter
          const encodedSearchTerm = encodeURIComponent(searchTerm);
          const filterExpression = `startswith(displayName,'${encodedSearchTerm}') or startswith(userPrincipalName,'${encodedSearchTerm}') or startswith(mail,'${encodedSearchTerm}')`;
          apiCall = apiCall.filter(filterExpression);
        }

        const response = await apiCall.select(USER_BASE_FIELDS).top(top).get();

        if (!response?.value?.length) {
          return [];
        }

        // Process users with photos
        const users = await Promise.allSettled(
          response.value.map(
            async (user: {
              id: string;
              displayName: string;
              jobTitle?: string;
              department?: string;
              mail?: string;
              userPrincipalName: string;
            }): Promise<IDirectReport> => {
              const userInfo: IDirectReport = {
                id: user.id,
                displayName: user.displayName,
                jobTitle: user.jobTitle,
                department: user.department,
                mail: user.mail,
                userPrincipalName: user.userPrincipalName,
              };

              // Fetch photo separately (non-blocking)
              try {
                userInfo.photoUrl = `${
                  context.pageContext.web.absoluteUrl
                }//_layouts/15/userphoto.aspx?size=L&username=${
                  user.userPrincipalName || ""
                }`;
              } catch {
                // Photo fetch failure is not critical
              }

              return userInfo;
            }
          )
        );

        return users
          .filter(
            (result: PromiseSettledResult<IDirectReport>) =>
              result.status === "fulfilled"
          )
          .map(
            (result: PromiseSettledResult<IDirectReport>) =>
              (result as PromiseFulfilledResult<IDirectReport>).value
          );
      } catch (error) {
        throw handleGraphError(error, "searchUsers", { searchTerm, top });
      }
    },
    []
  );

  /**
   * Check if a user has direct reports (lightweight check)
   */
  const hasDirectReports = React.useCallback(
    async (userId: string): Promise<boolean> => {
      try {
        const client = await getGraphClient();

        const response = await client
          .api(`/users/${userId}/directReports`)
          .top(1)
          .get();

        return Boolean(response?.value?.length);
      } catch {
        return false;
      }
    },
    []
  );

  /**
   * Get users with pagination support
   */
  const getUsers = React.useCallback(
    async (
      pageToken?: string,
      pageSize: number = 20,
      filter?: string,
      sortColumn?: string,
      sortDirection: "ascending" | "descending" = "ascending"
    ): Promise<IUsersPageResultSimplified> => {
      try {
        const client = await getGraphClient();
        let response: {
          value?: Array<{
            "@removed"?: boolean;
            id?: string;
            displayName?: string;
            mail?: string;
            userPrincipalName?: string;
            jobTitle?: string;
            department?: string;
            officeLocation?: string;
            businessPhones?: string[];
            mobilePhone?: string;
            companyName?: string;
            employeeId?: string;
            managerId?: string;
            userType?: string;
          }>;
          "@odata.nextLink"?: string;
          "@odata.count"?: number;
        };

        if (pageToken) {
          // Use next page token
          response = await client.api(pageToken).get();
        } else {
          // Initial request
          let apiCall = client
            .api("/users")
            .select(`${USER_BASE_FIELDS},managerId`)
            .top(pageSize)
            .count(true);

          // Add sorting if specified and supported
          if (sortColumn) {
            // Validate that the sort column is supported by Microsoft Graph API
            if (!SORTABLE_USER_PROPERTIES.includes(sortColumn)) {
              logError(
                "useGraphAPIs",
                `Sort property '${sortColumn}' is not supported by Microsoft Graph API for User objects. Skipping server-side sorting.`,
                new Error(`Unsupported sort property: ${sortColumn}`),
                ErrorType.SYSTEM,
                sanitizeUserData({
                  sortColumn,
                  operation: "getUsers_validateSortColumn",
                })
              );
            } else {
              const direction = sortDirection === "descending" ? "desc" : "asc";
              apiCall = apiCall.orderby(`${sortColumn} ${direction}`);
            }
          }

          if (filter) {
            apiCall = apiCall.header("ConsistencyLevel", "eventual");

            // Parse and apply filter
            if (filter.includes("$")) {
              // Query string format
              const params = new URLSearchParams(
                filter.startsWith("?") ? filter.slice(1) : filter
              );

              const searchParam = params.get("$search");
              if (searchParam) {
                const searchTerm = searchParam.replace(/^"|"$/g, "");
                if (searchTerm.includes(":")) {
                  apiCall = apiCall.search(`"${searchTerm}"`);
                } else {
                  const filterExpression = `startswith(displayName,'${searchTerm}') or startswith(mail,'${searchTerm}') or contains(displayName,'${searchTerm}')`;
                  apiCall = apiCall.filter(filterExpression);
                }
              }

              const filterParam = params.get("$filter");
              if (filterParam) {
                // Encode special characters in filter values to prevent OData issues
                const encodedFilter = encodeURIComponent(filterParam);
                apiCall = apiCall.filter(encodedFilter);
              }

              const orderByParam = params.get("$orderby");
              if (orderByParam) {
                apiCall = apiCall.orderby(orderByParam);
              }
            } else if (
              filter.includes(" eq ") ||
              filter.includes("contains(") ||
              filter.includes("startswith(")
            ) {
              // Direct OData filter - encode special characters
              const encodedFilter = encodeURIComponent(filter);
              apiCall = apiCall.filter(encodedFilter);
            } else {
              // Simple search term - encode special characters
              const encodedSearchTerm = encodeURIComponent(filter);
              const filterExpression = `startswith(displayName,'${encodedSearchTerm}') or contains(displayName,'${encodedSearchTerm}') or startswith(mail,'${encodedSearchTerm}')`;
              apiCall = apiCall.filter(filterExpression);
            }
          }

          response = await apiCall.get();
        }

        const users: IUserProfile[] =
          response?.value
            ?.filter((user: { "@removed"?: boolean }) => !user["@removed"])
            ?.map(
              (user: {
                id?: string;
                displayName?: string;
                mail?: string;
                userPrincipalName?: string;
                jobTitle?: string;
                department?: string;
                officeLocation?: string;
                businessPhones?: string[];
                mobilePhone?: string;
                companyName?: string;
                employeeId?: string;
                managerId?: string;
                userType?: string;
              }): IUserProfile => ({
                id: user.id || "",
                displayName: user.displayName || "",
                mail: user.mail || user.userPrincipalName || "",
                userPrincipalName: user.userPrincipalName || "",
                jobTitle: user.jobTitle,
                department: user.department,
                officeLocation: user.officeLocation,
                businessPhones: user.businessPhones,
                mobilePhone: user.mobilePhone,
                companyName: user.companyName,
                employeeId: user.employeeId,
                managerId: user.managerId,
                aboutMe: undefined,
                skills: undefined,
                photoUrl: undefined,
                userType: user.userType || "",
              })
            ) || [];

        return {
          users,
          hasMore: Boolean(response?.["@odata.nextLink"]),
          nextPageToken: response?.["@odata.nextLink"],
          totalCount: response?.["@odata.count"],
        };
      } catch (error) {
        throw handleGraphError(error, "getUsers", {
          pageToken: pageToken ? "present" : "undefined",
          pageSize,
          filter: filter ? filter.substring(0, 100) : "undefined",
        });
      }
    },
    [context]
  );

  // Early return with error functions if context is not available
  if (!context) {
    const errorFn = async (): Promise<never> => {
      throw new Error("Context not available");
    };

    return {
      getUserProfile: errorFn,
      getUserManager: errorFn,
      getAllManagers: errorFn,
      getUserDirectReports: errorFn,
      getUserPeers: errorFn,
      getUserPhoto: errorFn,
      searchUsers: errorFn,
      hasDirectReports: errorFn,
      getTotalDirectReports: errorFn,
      getUsers: errorFn,
    } as IUseGraphAPIsResult;
  }

  return {
    getUserProfile,
    getUserManager,
    getAllManagers,
    getUserDirectReports,
    getUserPeers,
    getUserPhoto,
    searchUsers,
    hasDirectReports,
    getTotalDirectReports,
    getUsers,
  };
};
