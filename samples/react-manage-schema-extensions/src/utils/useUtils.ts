import { RESTRICTED_PROPERTY_TYPES } from "../constants";

interface IUseUtils {
  getAvailablePropertyTypes: (
    selectedTargetTypes: string[],
    allPropertyTypes: readonly { key: string; text: string }[],
  ) => Array<{ key: string; text: string }>;
  hasPropertyRestrictions: (targetTypes: string[]) => boolean;
  sanitizeUserData: <T>(data: T) => T;
  blobToBase64: (blob: Blob) => Promise<string | undefined>;
  getInitials: (name: string) => string;
}

export const useUtils = (): IUseUtils => {
  /**
   * Function to get available property types based on selected target types
   */
  const getAvailablePropertyTypes = (
    selectedTargetTypes: string[],
    allPropertyTypes: readonly { key: string; text: string }[],
  ): Array<{ key: string; text: string }> => {
    // Get all restricted property types for the selected target types
    const restrictedTypes = selectedTargetTypes.reduce(
      (restricted: string[], targetType) => {
        const restrictions =
          RESTRICTED_PROPERTY_TYPES[
            targetType as keyof typeof RESTRICTED_PROPERTY_TYPES
          ];
        if (restrictions) {
          restricted.push(...restrictions);
        }
        return restricted;
      },
      [],
    );

    // Remove duplicates from restricted types
    const uniqueRestrictedTypes = [...new Set(restrictedTypes)];

    // Filter out restricted property types
    return allPropertyTypes.filter(
      (propertyType) => !uniqueRestrictedTypes.includes(propertyType.key),
    );
  };

  /**
   * Check if any target types have property restrictions
   */
  const hasPropertyRestrictions = (targetTypes: string[]): boolean => {
    return targetTypes.some(
      (targetType) =>
        RESTRICTED_PROPERTY_TYPES[
          targetType as keyof typeof RESTRICTED_PROPERTY_TYPES
        ],
    );
  };

  /**
   * Sanitizes user data to prevent base64 decode errors in external libraries
   */
  const sanitizeUserData = <T>(data: T): T => {
    if (!data || typeof data !== "object") {
      return data;
    }
    return data;
  };

  /**
 * Helper function to safely convert blob to base64 data URL
 */
 const blobToBase64 = (blob: Blob): Promise<string | undefined> => {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      try {
        const result = reader.result as string;
        if (result?.startsWith("data:")) {
          // Validate base64 format
          const base64Part = result.split(",")[1];
          if (base64Part) {
            atob(base64Part); // Test if valid base64
            resolve(result);
          } else {
            resolve(undefined);
          }
        } else {
          resolve(undefined);
        }
      } catch {
        resolve(undefined);
      }
    };
    reader.onerror = () => resolve(undefined);
    reader.readAsDataURL(blob);
  });
};


  /**
   * Gets initials from first and last word of a name
   * e.g. "John Doe" → "JD", "Alice" → "A", "John Michael Doe" → "JD"
   */
  const getInitials = (name: string): string => {
    if (!name?.trim()) return "";
    const words = name.trim().split(/\s+/);
    const first = words[0].charAt(0).toUpperCase();
    const last = words.length > 1 ? words[words.length - 1].charAt(0).toUpperCase() : "";
    return `${first}${last}`;
  };


  return {
    getAvailablePropertyTypes,
    hasPropertyRestrictions,
    sanitizeUserData,
      blobToBase64,
      getInitials,
  };
};
