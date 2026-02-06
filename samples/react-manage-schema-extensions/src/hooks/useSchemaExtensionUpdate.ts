import * as React from "react";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ErrorType, useLogging } from "@spteck/m365-hooks";
import { useUtils } from "../utils/useUtils";
 

export interface ISchemaExtensionUpdateProps {
  context: BaseComponentContext;
}

export interface IUserSchemaExtensionData {
  skillsline1?: string;
  skillsline2?: string;
  skillsline3?: string;
  skillsline4?: string;
  skillsline5?: string;
  projectsline1?: string;
  projectsline2?: string;
  projectsline3?: string;
  projectsline4?: string;
  projectsline5?: string;
}

export interface ISchemaExtensionUpdateResult {
  getUserSchemaExtension: (userId: string) => Promise<IUserSchemaExtensionData | undefined>;
  updateUserSchemaExtension: (userId: string, data: IUserSchemaExtensionData) => Promise<boolean>;
  isLoading: boolean;
  error: string | undefined;
}

const SCHEMA_EXTENSION_ID = "extinwbntrw_spUserProfile";

export const useSchemaExtensionUpdate = ({
  context,
}: ISchemaExtensionUpdateProps): ISchemaExtensionUpdateResult => {
  const { logError } = useLogging();
  const { sanitizeUserData } = useUtils();
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  // Create Graph client function
  const getGraphClient = React.useCallback(async (): Promise<MSGraphClientV3> => {
    if (!context) {
      throw new Error("Context not available");
    }
    return await context.msGraphClientFactory.getClient("3");
  }, [context]);

  /**
   * Helper function to handle Graph API errors consistently
   */
  const handleGraphError = React.useCallback(
    (error: unknown, operation: string, additionalContext?: Record<string, unknown>): Error => {
      const graphError = error as Error & { code?: string; status?: number };

      console.error(`Graph API Error in ${operation}:`, error);

      logError(
        "useSchemaExtensionUpdate",
        `Error in ${operation}`,
        graphError,
        ErrorType.SYSTEM,
        sanitizeUserData({
          operation,
          errorCode: graphError.code,
          statusCode: graphError.status,
          ...additionalContext,
        })
      );

      // Create user-friendly error message
      let errorMessage = `Failed to ${operation.toLowerCase().replace(/([A-Z])/g, ' $1')}`;
      
      if (graphError.code === "Forbidden" || graphError.status === 403) {
        errorMessage = "You don't have permission to perform this operation";
      } else if (graphError.code === "NotFound" || graphError.status === 404) {
        errorMessage = "User not found";
      } else if (graphError.code === "BadRequest" || graphError.status === 400) {
        errorMessage = "Invalid request. Please check the data provided";
      } else if (graphError.code === "TooManyRequests" || graphError.status === 429) {
        errorMessage = "Too many requests. Please try again later";
      }

      return new Error(errorMessage);
    },
    [logError, sanitizeUserData]
  );

  /**
   * Get user's schema extension data
   */
  const getUserSchemaExtension = React.useCallback(
    async (userId: string): Promise<IUserSchemaExtensionData | undefined> => {
      if (!userId?.trim()) {
        throw new Error("User ID is required");
      }

      setIsLoading(true);
      setError(undefined);

      try {
        const client = await getGraphClient();
        
        const response = await client
          .api(`/users/${userId}`)
          .select(`${SCHEMA_EXTENSION_ID}`)
          .get();

        const extensionData = response[SCHEMA_EXTENSION_ID];
        
        if (!extensionData) {
          return {
            skillsline1: "",
            skillsline2: "",
            skillsline3: "",
            skillsline4: "",
            skillsline5: "",
            projectsline1: "",
            projectsline2: "",
            projectsline3: "",
            projectsline4: "",
            projectsline5: "",
          };
        }

        return {
          skillsline1: extensionData.skillsline1 || "",
          skillsline2: extensionData.skillsline2 || "",
          skillsline3: extensionData.skillsline3 || "",
          skillsline4: extensionData.skillsline4 || "",
          skillsline5: extensionData.skillsline5 || "",
          projectsline1: extensionData.projectsline1 || "",
          projectsline2: extensionData.projectsline2 || "",
          projectsline3: extensionData.projectsline3 || "",
          projectsline4: extensionData.projectsline4 || "",
          projectsline5: extensionData.projectsline5 || "",
        };
      } catch (error: unknown) {
        console.error("Error fetching user schema extension:", error);
        const graphError = error as Error & { code?: string; status?: number };
        
        // Handle specific Graph API errors without causing loops
        if (graphError.code === "Request_ResourceNotFound" || graphError.status === 404) {
          // User doesn't have extension data - return empty data
          return {
            skillsline1: "",
            skillsline2: "",
            skillsline3: "",
            skillsline4: "",
            skillsline5: "",
            projectsline1: "",
            projectsline2: "",
            projectsline3: "",
            projectsline4: "",
            projectsline5: "",
          };
        }
        
        const handledError = handleGraphError(error, "getUserSchemaExtension", { userId });
        setError(handledError.message);
        throw handledError;
      } finally {
        setIsLoading(false);
      }
    },
    [getGraphClient]
  );

  /**
   * Update user's schema extension data
   */
  const updateUserSchemaExtension = React.useCallback(
    async (userId: string, data: IUserSchemaExtensionData): Promise<boolean> => {
      if (!userId?.trim()) {
        throw new Error("User ID is required");
      }

      setIsLoading(true);
      setError(undefined);

      try {
        const client = await getGraphClient();
        
        // Prepare the update payload
        const updatePayload = {
          [SCHEMA_EXTENSION_ID]: {
            skillsline1: data.skillsline1 || null,
            skillsline2: data.skillsline2 || null,
            skillsline3: data.skillsline3 || null,
            skillsline4: data.skillsline4 || null,
            skillsline5: data.skillsline5 || null,
            projectsline1: data.projectsline1 || null,
            projectsline2: data.projectsline2 || null,
            projectsline3: data.projectsline3 || null,
            projectsline4: data.projectsline4 || null,
            projectsline5: data.projectsline5 || null,
          }
        };

        await client
          .api(`/users/${userId}`)
          .patch(updatePayload);

        return true;
      } catch (error: unknown) {
        console.error("Error updating user schema extension:", error);
        const handledError = handleGraphError(error, "updateUserSchemaExtension", { 
          userId,
          dataKeys: Object.keys(data)
        });
        setError(handledError.message);
        throw handledError;
      } finally {
        setIsLoading(false);
      }
    },
    [getGraphClient]
  );

  return {
    getUserSchemaExtension,
    updateUserSchemaExtension,
    isLoading,
    error,
  };
};