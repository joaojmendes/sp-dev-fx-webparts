import * as React from "react";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ErrorType, useLogging } from "@spteck/m365-hooks";
import { ISchemaExtension, ISchemaExtensionProperty } from "../models/ISchemaExtension";
import { ETargetTypes } from "../constants";
import { useSchemaExtensionStorage } from "./useSchemaExtensionStorage";
import { ISchemaTenantProperty } from "../models/ISchemaTeantProperty";

export interface IResourceSchemaExtensionProps {
  context: BaseComponentContext;
}

/**
 * Represents a generic resource from Microsoft Graph
 */
export interface IGraphResource {
  id: string;
  displayName?: string;
  userPrincipalName?: string;
  mail?: string;
  description?: string;
  name?: string;
  subject?: string;
  [key: string]: unknown;
}

/**
 * Resource attribute data - dynamic based on schema properties
 */
export interface IResourceAttributeData {
  [key: string]: string | number | boolean | undefined;
}

/**
 * Schema extension with parsed properties for form rendering
 */
export interface ISchemaExtensionWithProperties extends ISchemaExtension {
  parsedProperties: ISchemaExtensionProperty[];
}

export interface IResourceSchemaExtensionResult {
  // Schema operations
  getSchemaExtensions: () => Promise<ISchemaExtension[]>;
  getSchemaExtension: (id: string) => Promise<ISchemaExtension | undefined>;
  
  // Resource operations
  getResourcesByType: (targetType: string, searchQuery?: string) => Promise<IGraphResource[]>;
  getResourceAttributes: (targetType: string, resourceId: string, schemaExtensionId: string) => Promise<IResourceAttributeData>;
  updateResourceAttributes: (
    targetType: string, 
    resourceId: string, 
    schemaExtensionId: string, 
    data: IResourceAttributeData
  ) => Promise<boolean>;
  
  // State
  isLoading: boolean;
  error: string | undefined;
}

/**
 * Maps target types to Microsoft Graph API endpoints
 */
const TARGET_TYPE_ENDPOINTS: Record<string, string> = {
  [ETargetTypes.User]: "/users",
  [ETargetTypes.Group]: "/groups",
  [ETargetTypes.Device]: "/devices",
  [ETargetTypes.Event]: "/me/events",
  [ETargetTypes.Message]: "/me/messages",
  [ETargetTypes.Organization]: "/organization",
  [ETargetTypes.Contact]: "/contacts",
  [ETargetTypes.AdministrativeUnit]: "/directory/administrativeUnits",
  [ETargetTypes.Post]: "/groups/{id}/threads/{id}/posts",
  [ETargetTypes.TodoTask]: "/me/todo/lists/{listId}/tasks",
  [ETargetTypes.TodoTaskList]: "/me/todo/lists",
};

/**
 * Maps target types to their display name property
 */
const DISPLAY_PROPERTY_MAP: Record<string, string> = {
  [ETargetTypes.User]: "displayName",
  [ETargetTypes.Group]: "displayName",
  [ETargetTypes.Device]: "displayName",
  [ETargetTypes.Event]: "subject",
  [ETargetTypes.Message]: "subject",
  [ETargetTypes.Organization]: "displayName",
  [ETargetTypes.Contact]: "displayName",
  [ETargetTypes.AdministrativeUnit]: "displayName",
  [ETargetTypes.Post]: "body",
  [ETargetTypes.TodoTask]: "title",
  [ETargetTypes.TodoTaskList]: "displayName",
};

/**
 * Hook for managing schema extension attributes on resources
 */
export const useResourceSchemaExtension = ({
  context,
}: IResourceSchemaExtensionProps): IResourceSchemaExtensionResult => {
  const { logError, logInfo, logWarning } = useLogging();
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  // Initialize storage hook for schema extension configuration (uses internal list)
  const { getProperty, getAppCatalogUrl } = useSchemaExtensionStorage(context);

  // Tenant property key (same as useSchemaExtension)
  const TENANT_PROPERTY_KEY = "ManageSchemaExtensionsCreatedSchemaIds";

  // Create Graph client function
  const getGraphClient = React.useCallback(async (): Promise<MSGraphClientV3> => {
    if (!context) {
      throw new Error("Context not available");
    }
    return await context.msGraphClientFactory.getClient("3");
  }, [context]);

  /**
   * Get created schema extension IDs from SharePoint list storage
   */
  const getCreatedSchemaIds = React.useCallback(async (): Promise<ISchemaTenantProperty[]> => {
    try {
      const appCatalogUrl = await getAppCatalogUrl();
      if (!appCatalogUrl) {
        throw new Error("App catalog URL not available");
      }
      const property = await getProperty(TENANT_PROPERTY_KEY);
      if (!property?.Value) {
        return [];
      }
      try {
        return JSON.parse(property.Value) || [];
      } catch (parseError) {
        logWarning(
          "getCreatedSchemaIds",
          "Failed to parse stored schema IDs, returning empty array:",
          JSON.stringify(parseError)
        );
        return [];
      }
    } catch (err) {
      logError(
        "getCreatedSchemaIds",
        "Failed to get created schema IDs from storage:",
        JSON.stringify(err)
      );
      return [];
    }
  }, [getAppCatalogUrl, getProperty, logError, logWarning]);

  /**
   * Helper function to handle Graph API errors consistently
   */
  const handleGraphError = React.useCallback(
    (graphError: unknown, operation: string, additionalContext?: Record<string, unknown>): Error => {
      const err = graphError as Error & { code?: string; status?: number };

      console.error(`Graph API Error in ${operation}:`, graphError);

      logError(
        "useResourceSchemaExtension",
        `Error in ${operation}`,
        err,
        ErrorType.SYSTEM,
        {
          operation,
          errorCode: err.code,
          statusCode: err.status,
          ...additionalContext,
        }
      );

      // Create user-friendly error message
      let errorMessage = `Failed to ${operation.toLowerCase().replace(/([A-Z])/g, " $1")}`;

      if (err.code === "Forbidden" || err.status === 403) {
        errorMessage = "You don't have permission to perform this operation";
      } else if (err.code === "NotFound" || err.status === 404) {
        errorMessage = "Resource not found";
      } else if (err.code === "BadRequest" || err.status === 400) {
        errorMessage = "Invalid request. Please check the data provided";
      } else if (err.code === "TooManyRequests" || err.status === 429) {
        errorMessage = "Too many requests. Please try again later";
      }

      return new Error(errorMessage);
    },
    [logError]
  );

  /**
   * Get all available schema extensions from the internal list
   * This uses the same approach as useSchemaExtension - get IDs from storage, then fetch details from Graph
   */
  const getSchemaExtensions = React.useCallback(async (): Promise<ISchemaExtension[]> => {
    setIsLoading(true);
    setError(undefined);

    try {
      // Get schema IDs from the internal SharePoint list
      const storedIds = await getCreatedSchemaIds();
      
      if (storedIds.length === 0) {
        logInfo("getSchemaExtensions", "No schema extension IDs found in storage");
        return [];
      }

      logInfo("getSchemaExtensions", `Found ${storedIds.length} schema extension IDs in storage`);

      // Fetch each schema extension by ID from Graph API
      const client = await getGraphClient();
      const schemaExtensions: ISchemaExtension[] = [];

      for (const schema of storedIds) {
        try {
          const response = await client
            .api(`/schemaExtensions/${schema.schemaId}`)
            .get();
          
          if (response) {
            schemaExtensions.push(response as ISchemaExtension);
          }
        } catch (schemaError) {
          // Log warning but continue with other schemas
          logWarning(
            "getSchemaExtensions",
            `Failed to fetch schema extension ${schema.schemaId}:`,
            JSON.stringify(schemaError)
          );
        }
      }

      logInfo("getSchemaExtensions", `Retrieved ${schemaExtensions.length} schema extensions from Graph`);
      
      return schemaExtensions;
    } catch (err: unknown) {
      const handledError = handleGraphError(err, "getSchemaExtensions");
      setError(handledError.message);
      throw handledError;
    } finally {
      setIsLoading(false);
    }
  }, [getGraphClient, getCreatedSchemaIds, handleGraphError, logInfo, logWarning]);

  /**
   * Get a specific schema extension by ID
   */
  const getSchemaExtension = React.useCallback(
    async (id: string): Promise<ISchemaExtension | undefined> => {
      if (!id?.trim()) {
        throw new Error("Schema extension ID is required");
      }

      setIsLoading(true);
      setError(undefined);

      try {
        const client = await getGraphClient();
        
        const response = await client
          .api(`/schemaExtensions/${id}`)
          .get();

        logInfo("getSchemaExtension", `Retrieved schema extension: ${id}`);
        
        return response;
      } catch (err: unknown) {
        const graphError = err as Error & { code?: string; status?: number };
        
        if (graphError.code === "Request_ResourceNotFound" || graphError.status === 404) {
          return undefined;
        }
        
        const handledError = handleGraphError(err, "getSchemaExtension", { schemaId: id });
        setError(handledError.message);
        throw handledError;
      } finally {
        setIsLoading(false);
      }
    },
    [getGraphClient, handleGraphError, logInfo]
  );

  /**
   * Get resources by target type with optional search
   */
  const getResourcesByType = React.useCallback(
    async (targetType: string, searchQuery?: string): Promise<IGraphResource[]> => {
      if (!targetType?.trim()) {
        throw new Error("Target type is required");
      }

      const endpoint = TARGET_TYPE_ENDPOINTS[targetType];
      if (!endpoint) {
        throw new Error(`Unsupported target type: ${targetType}`);
      }

      setIsLoading(true);
      setError(undefined);

      try {
        const client = await getGraphClient();
        const displayProperty = DISPLAY_PROPERTY_MAP[targetType] || "displayName";

        let apiRequest = client
          .api(endpoint)
          .top(50);

        // Add search filter if provided
        if (searchQuery?.trim()) {
          // Use $search for users and groups (more powerful search)
          if (targetType === ETargetTypes.User || targetType === ETargetTypes.Group) {
            apiRequest = apiRequest
              .header("ConsistencyLevel", "eventual")
              .search(`"displayName:${searchQuery}"`);
          } else {
            // Use startsWith filter for other resource types
            apiRequest = apiRequest.filter(`startswith(${displayProperty},'${searchQuery}')`);
          }
        }

        // Select relevant properties based on type
        switch (targetType) {
          case ETargetTypes.User:
            apiRequest = apiRequest.select("id,displayName,userPrincipalName,mail,jobTitle");
            break;
          case ETargetTypes.Group:
            apiRequest = apiRequest.select("id,displayName,description,mail,groupTypes");
            break;
          case ETargetTypes.Device:
            apiRequest = apiRequest.select("id,displayName,operatingSystem,operatingSystemVersion");
            break;
          case ETargetTypes.Event:
            apiRequest = apiRequest.select("id,subject,start,end,organizer");
            break;
          case ETargetTypes.Message:
            apiRequest = apiRequest.select("id,subject,from,receivedDateTime");
            break;
          case ETargetTypes.Organization:
            apiRequest = apiRequest.select("id,displayName,verifiedDomains");
            break;
          case ETargetTypes.Contact:
            apiRequest = apiRequest.select("id,displayName,emailAddresses");
            break;
          case ETargetTypes.AdministrativeUnit:
            apiRequest = apiRequest.select("id,displayName,description");
            break;
          default:
            apiRequest = apiRequest.select("id,displayName");
        }

        const response = await apiRequest.get();

        logInfo("getResourcesByType", `Retrieved ${response.value?.length || 0} ${targetType} resources`);
        
        return response.value || [];
      } catch (err: unknown) {
        const handledError = handleGraphError(err, "getResourcesByType", { targetType, searchQuery });
        setError(handledError.message);
        throw handledError;
      } finally {
        setIsLoading(false);
      }
    },
    [getGraphClient, handleGraphError, logInfo]
  );

  /**
   * Get resource's schema extension attributes
   */
  const getResourceAttributes = React.useCallback(
    async (
      targetType: string,
      resourceId: string,
      schemaExtensionId: string
    ): Promise<IResourceAttributeData> => {
      if (!targetType?.trim() || !resourceId?.trim() || !schemaExtensionId?.trim()) {
        throw new Error("Target type, resource ID, and schema extension ID are required");
      }

      const endpoint = TARGET_TYPE_ENDPOINTS[targetType];
      if (!endpoint) {
        throw new Error(`Unsupported target type: ${targetType}`);
      }

      setIsLoading(true);
      setError(undefined);

      try {
        const client = await getGraphClient();
        
        // Request the resource with the schema extension data
        // The schema extension ID is used in $select to get the extension properties
        const response = await client
          .api(`${endpoint}/${resourceId}`)
          .select(`id,${schemaExtensionId}`)
          .get();

        logInfo("getResourceAttributes", `Response for ${targetType}/${resourceId}`, { response: JSON.stringify(response) });

        // The extension data is stored under the schema extension ID key
        // Schema extension IDs are in format: ext{appId}_{schemaName}
        const extensionData = response[schemaExtensionId];

        if (extensionData) {
          logInfo("getResourceAttributes", `Found existing attributes for ${schemaExtensionId}`, { extensionData: JSON.stringify(extensionData) });
        } else {
          logInfo("getResourceAttributes", `No existing attributes found for ${schemaExtensionId} on ${targetType}/${resourceId}`);
        }
        
        return extensionData || {};
      } catch (err: unknown) {
        const graphError = err as Error & { code?: string; status?: number };
        
        // If resource not found or no extension data, return empty object
        if (graphError.code === "Request_ResourceNotFound" || graphError.status === 404) {
          logInfo("getResourceAttributes", `Resource not found: ${targetType}/${resourceId}`);
          return {};
        }

        // If the extension property doesn't exist on the resource, Graph might return an error
        // In this case, we should return empty object (no data set yet)
        if (graphError.message?.includes("Could not find a property named")) {
          logInfo("getResourceAttributes", `Schema extension ${schemaExtensionId} not yet applied to ${targetType}/${resourceId}`);
          return {};
        }
        
        const handledError = handleGraphError(err, "getResourceAttributes", {
          targetType,
          resourceId,
          schemaExtensionId,
        });
        setError(handledError.message);
        throw handledError;
      } finally {
        setIsLoading(false);
      }
    },
    [getGraphClient, handleGraphError, logInfo]
  );

  /**
   * Update resource's schema extension attributes
   */
  const updateResourceAttributes = React.useCallback(
    async (
      targetType: string,
      resourceId: string,
      schemaExtensionId: string,
      data: IResourceAttributeData
    ): Promise<boolean> => {
      if (!targetType?.trim() || !resourceId?.trim() || !schemaExtensionId?.trim()) {
        throw new Error("Target type, resource ID, and schema extension ID are required");
      }

      const endpoint = TARGET_TYPE_ENDPOINTS[targetType];
      if (!endpoint) {
        throw new Error(`Unsupported target type: ${targetType}`);
      }

      setIsLoading(true);
      setError(undefined);

      try {
        const client = await getGraphClient();
        
        // Build the update payload with schema extension data
        const updatePayload = {
          [schemaExtensionId]: data,
        };

        await client
          .api(`${endpoint}/${resourceId}`)
          .patch(updatePayload);

        logInfo("updateResourceAttributes", `Updated attributes for ${targetType}/${resourceId}`);
        
        return true;
      } catch (err: unknown) {
        const handledError = handleGraphError(err, "updateResourceAttributes", {
          targetType,
          resourceId,
          schemaExtensionId,
        });
        setError(handledError.message);
        throw handledError;
      } finally {
        setIsLoading(false);
      }
    },
    [getGraphClient, handleGraphError, logInfo]
  );

  return {
    getSchemaExtensions,
    getSchemaExtension,
    getResourcesByType,
    getResourceAttributes,
    updateResourceAttributes,
    isLoading,
    error,
  };
};
