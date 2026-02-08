import * as React from "react";
import {
  Tag,
  TagPicker,
  TagPickerControl,
  TagPickerGroup,
  TagPickerInput,
  TagPickerList,
  TagPickerOption,
  TagPickerProps,
  Text,
  Avatar,
} from "@fluentui/react-components";
import { useCallback, useEffect, useRef, useState } from "react";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { IGraphResource, IResourcePageResult, useResourceSchemaExtension } from "../../hooks/useResourceSchemaExtension";
import { ETargetTypes } from "../../constants";
 
import {
  Person24Regular,
  PeopleTeam24Regular,
  Phone24Regular,
  Calendar24Regular,
  Mail24Regular,
  Organization24Regular,
  ContactCard24Regular,
  Shield24Regular,
} from "@fluentui/react-icons";
import { useResourcePickerStyles } from "./useResourcePickerStyles";
import { useUtils } from "../../utils/useUtils";

export interface IResourcePickerProps {
  context: BaseComponentContext;
  targetType: string;
  selectedResources?: IGraphResource[];
  onSelectionChange?: (selectedResources: IGraphResource[]) => void;
  placeholder?: string;
  disabled?: boolean;
  maxSelectedOptions?: number;
  width?: string | number;
  className?: string;
}

/**
 * Get icon for target type
 */
const getTargetTypeIcon = (targetType: string): JSX.Element => {
  switch (targetType) {
    case ETargetTypes.User:
      return <Person24Regular />;
    case ETargetTypes.Group:
      return <PeopleTeam24Regular />;
    case ETargetTypes.Device:
      return <Phone24Regular />;
    case ETargetTypes.Event:
      return <Calendar24Regular />;
    case ETargetTypes.Message:
      return <Mail24Regular />;
    case ETargetTypes.Organization:
      return <Organization24Regular />;
    case ETargetTypes.Contact:
      return <ContactCard24Regular />;
    case ETargetTypes.AdministrativeUnit:
      return <Shield24Regular />;
    default:
      return <Person24Regular />;
  }
};

/**
 * Get display name for a resource based on its type
 */
const getResourceDisplayName = (resource: IGraphResource, targetType: string): string => {
  switch (targetType) {
    case ETargetTypes.User:
      return resource.displayName || resource.userPrincipalName || resource.id;
    case ETargetTypes.Group:
      return resource.displayName || resource.mail || resource.id;
    case ETargetTypes.Device:
      return resource.displayName || resource.id;
    case ETargetTypes.Event:
      return resource.subject || resource.id;
    case ETargetTypes.Message:
      return resource.subject || resource.id;
    default:
      return resource.displayName || resource.name || resource.id;
  }
};

/**
 * Get secondary text for a resource based on its type
 */
const getResourceSecondaryText = (resource: IGraphResource, targetType: string): string => {
  switch (targetType) {
    case ETargetTypes.User:
      return (resource.mail as string) || (resource.userPrincipalName as string) || "";
    case ETargetTypes.Group:
      return (resource.description as string) || (resource.mail as string) || "";
    case ETargetTypes.Device:
      return `${resource.operatingSystem || ""} ${resource.operatingSystemVersion || ""}`.trim();
    case ETargetTypes.Event:
      return resource.start ? `Starts: ${new Date(resource.start as string).toLocaleString()}` : "";
    case ETargetTypes.Message:
      return resource.receivedDateTime ? `Received: ${new Date(resource.receivedDateTime as string).toLocaleString()}` : "";
    default:
      return (resource.description as string) || "";
  }
};

export const ResourcePicker: React.FunctionComponent<IResourcePickerProps> = ({
  context,
  targetType,
  selectedResources = [],
  onSelectionChange,
  placeholder,
  disabled = false,
  maxSelectedOptions = 1,
  width,
  className,
}) => { 
  const styles = useResourcePickerStyles();
  const { getInitials } = useUtils();
  const { getResourcesByType } = useResourceSchemaExtension({ context });
  
  const [availableResources, setAvailableResources] = useState<IGraphResource[]>([]);
  const [isSearching, setIsSearching] = useState(false);
  const [searchValue, setSearchValue] = useState("");
  const [hasLoaded, setHasLoaded] = useState(false);

  // Infinite scroll state
  const [hasMore, setHasMore] = useState(true);
  const [isLoadingMore, setIsLoadingMore] = useState(false);
  const [nextPageToken, setNextPageToken] = useState<string | undefined>(undefined);

  // Debounce timer ref
  const searchTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const listRef = useRef<HTMLDivElement>(null);

  // Convert selected resources to array of IDs for TagPicker
  const selectedResourceIds = selectedResources.map(resource => resource.id);

  // Get placeholder text based on target type
  const getPlaceholder = (): string => {
    if (placeholder) return placeholder;
    if (!targetType) return "Select a target type first...";
    return `Search for ${targetType.toLowerCase()}s...`;
  };

  /**
   * Load initial resources when target type changes
   */
  useEffect(() => {
    let isCancelled = false;

    const loadInitialResources = async (): Promise<void> => {
      if (!targetType) {
        setAvailableResources([]);
        setHasLoaded(false);
        setHasMore(false);
        setNextPageToken(undefined);
        return;
      }

      setIsSearching(true);
      try {
        const result = await getResourcesByType(targetType, undefined, undefined, 20);
        if (!isCancelled) {
          setAvailableResources(result.resources);
          setHasMore(result.hasMore);
          setNextPageToken(result.nextPageToken);
          setHasLoaded(true);
        }
      } catch (error) {
        console.error("Error loading initial resources:", error);
        if (!isCancelled) {
          setAvailableResources([]);
          setHasMore(false);
          setNextPageToken(undefined);
          setHasLoaded(true);
        }
      } finally {
        if (!isCancelled) {
          setIsSearching(false);
        }
      }
    };

    loadInitialResources().catch(console.error);

    return () => {
      isCancelled = true;
    };
  }, [targetType, getResourcesByType]);

  /**
   * Handle search input changes with debounce
   */
  const handleInputChange = useCallback(async (value: string) => {
    setSearchValue(value);

    if (!targetType) return;

    // Clear existing timer
    if (searchTimerRef.current) {
      clearTimeout(searchTimerRef.current);
    }

    if (!value.trim()) {
      // If empty, reload initial resources with pagination
      setIsSearching(true);
      try {
        const result = await getResourcesByType(targetType, undefined, undefined, 20);
        setAvailableResources(result.resources);
        setHasMore(result.hasMore);
        setNextPageToken(result.nextPageToken);
      } catch (error) {
        console.error("Error loading resources:", error);
      } finally {
        setIsSearching(false);
      }
      return;
    }

    // Debounce search
    searchTimerRef.current = setTimeout(async () => {
      setIsSearching(true);
      try {
        const results = await getResourcesByType(targetType, value.trim(), undefined, 20);
        setAvailableResources(results.resources);
        // Disable infinite scroll during search
        setHasMore(false);
        setNextPageToken(undefined);
      } catch (error) {
        console.error("Error searching resources:", error);
        setAvailableResources([]);
        setHasMore(false);
        setNextPageToken(undefined);
      } finally {
        setIsSearching(false);
      }
    }, 300);
  }, [targetType, getResourcesByType]);

  /**
   * Load more resources for infinite scroll
   */
  const loadMoreResources = useCallback(async () => {
    if (!hasMore || isLoadingMore || !nextPageToken || searchValue.trim()) {
      return;
    }

    setIsLoadingMore(true);
    try {
      const result = await getResourcesByType(targetType, undefined, nextPageToken, 20);
      setAvailableResources(prev => [...prev, ...result.resources]);
      setHasMore(result.hasMore);
      setNextPageToken(result.nextPageToken);
    } catch (error) {
      console.error("Error loading more resources:", error);
    } finally {
      setIsLoadingMore(false);
    }
  }, [getResourcesByType, targetType, hasMore, isLoadingMore, nextPageToken, searchValue]);

  /**
   * Handle scroll event for infinite scroll
   */
  const handleScroll = useCallback((event: React.UIEvent<HTMLDivElement>) => {
    if (!hasMore || isLoadingMore || searchValue.trim()) {
      return;
    }

    const { scrollTop, scrollHeight, clientHeight } = event.currentTarget;
    const threshold = 100;
    const isNearBottom = scrollTop + clientHeight >= scrollHeight - threshold;

    if (isNearBottom) {
      loadMoreResources().catch(console.error);
    }
  }, [hasMore, isLoadingMore, searchValue, loadMoreResources]);

  /**
   * Handle option selection and dismissal - follows the UserPicker pattern
   */
  const onOptionSelect: TagPickerProps["onOptionSelect"] = useCallback((e, data) => {
    if (data.value === "no-options" || data.value === "loading" || data.value === "loading-more") {
      return;
    }

    // Use the selectedOptions from data to update the selection
    const newSelectedResourceIds = data.selectedOptions;
    const newSelectedResources = selectedResources.filter(resource =>
      newSelectedResourceIds.includes(resource.id)
    );

    // If a new option was selected (not just dismissed), add it
    if (newSelectedResourceIds.length > selectedResourceIds.length) {
      const newResourceId = newSelectedResourceIds.find((id: string) => !selectedResourceIds.includes(id));
      if (newResourceId) {
        const selectedResource = availableResources.find(resource => resource.id === newResourceId);
        if (selectedResource) {
          // Check max limit
          if (maxSelectedOptions && selectedResources.length >= maxSelectedOptions) {
            return;
          }
          newSelectedResources.push(selectedResource);
        }
      }

      // Clear search input and reset to initial resources after selection
      setSearchValue("");
      // Load initial resources to reset the list with pagination
      if (targetType) {
        getResourcesByType(targetType, undefined, undefined, 20).then((result: IResourcePageResult) => {
          setAvailableResources(result.resources);
          setHasMore(result.hasMore);
          setNextPageToken(result.nextPageToken);
        }).catch(console.error);
      }
    }

    onSelectionChange?.(newSelectedResources);
  }, [availableResources, selectedResources, selectedResourceIds, onSelectionChange, maxSelectedOptions, getResourcesByType, targetType]);

  // Filter available resources to exclude already selected ones
  const tagPickerOptions = availableResources.filter(
    resource => !selectedResourceIds.includes(resource.id)
  );

  // Cleanup on unmount
  useEffect(() => {
    return () => {
      if (searchTimerRef.current) {
        clearTimeout(searchTimerRef.current);
      }
    };
  }, []);

  return (
    <div
      className={`${styles.container} ${className || ""}`}
      {...(width && { style: { width } })}
    >
      <TagPicker
          onOptionSelect={onOptionSelect}
          selectedOptions={selectedResourceIds}
          disabled={disabled || !targetType}
        >
          <TagPickerControl className={styles.tagPickerControl}>
            <TagPickerGroup aria-label="Selected Resources">
              {selectedResources.map((resource) => (
                <Tag
                  key={resource.id}
                  shape="circular"
                  media={
                    <Avatar
                      name={getResourceDisplayName(resource, targetType)}
                      size={20}
                      icon={getTargetTypeIcon(targetType)}
                      color="colorful"
                      initials={getInitials(getResourceDisplayName(resource, targetType))}
                    />
                  }
                  value={resource.id}
                  dismissible
                  disabled={disabled}
                >
                  {getResourceDisplayName(resource, targetType)}
                </Tag>
              ))}
            </TagPickerGroup>
            <TagPickerInput
              className={styles.tagPickerInput}
              aria-label="Select Resources"
              placeholder={getPlaceholder()}
              value={searchValue}
              onChange={(e) => handleInputChange(e.target.value)}
              disabled={disabled || !targetType}
            />
          </TagPickerControl>

          <TagPickerList
            ref={listRef}
            onScroll={handleScroll}
            className={styles.pickerList}
          >
            {isSearching ? (
              <TagPickerOption value="loading">
                Searching...
              </TagPickerOption>
            ) : tagPickerOptions.length > 0 ? (
              <>
                {tagPickerOptions.map((resource) => {
                  const displayName = getResourceDisplayName(resource, targetType);
                  const secondaryText = getResourceSecondaryText(resource, targetType);

                  return (
                    <TagPickerOption
                      key={resource.id}
                      value={resource.id}
                      text={displayName}
                      media={
                        <Avatar
                          name={displayName}
                          size={36}
                          icon={getTargetTypeIcon(targetType)}
                          color="colorful"
                          initials={getInitials(displayName)}
                        />
                      }
                    >
                      <div className={styles.optionContent}>
                        <Text weight="semibold">{displayName}</Text>
                        {secondaryText && (
                          <Text size={200} className={styles.optionSecondary}>
                            {secondaryText}
                          </Text>
                        )}
                      </div>
                    </TagPickerOption>
                  );
                })}
                {isLoadingMore && !searchValue.trim() && (
                  <TagPickerOption value="loading-more">
                    Loading more...
                  </TagPickerOption>
                )}
              </>
            ) : (
              <TagPickerOption value="no-options">
                {!targetType
                  ? "Select a target type first"
                  : searchValue
                  ? `No ${targetType.toLowerCase()}s found`
                  : hasLoaded
                  ? `No ${targetType.toLowerCase()}s available`
                  : "Start typing to search..."}
              </TagPickerOption>
            )}
          </TagPickerList>
        </TagPicker>
    </div>
  );
};
