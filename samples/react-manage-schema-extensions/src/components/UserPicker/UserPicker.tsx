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
} from "@fluentui/react-components";
import { useCallback, useEffect, useRef, useState } from "react";

import { BaseComponentContext } from "@microsoft/sp-component-base";

import { UserCard } from "@spteck/react-controls";
import { useUserPickerStyles } from "./useUserPickerStyles";
import { IUserProfile } from "../../models/IUserData";
import { IUsersPageResultSimplified, useGraphAPIs } from "../../hooks/useGraphAPIs";

export interface IUserPickerProps {
  context: BaseComponentContext;
  selectedUsers?: IUserProfile[];
  onSelectionChange?: (selectedUsers: IUserProfile[]) => void;
  placeholder?: string;
  disabled?: boolean;
  maxSelectedOptions?: number;
  width?: string | number;
  className?: string;
}

export const UserPicker: React.FunctionComponent<IUserPickerProps> = ({
  context,
  selectedUsers = [],
  onSelectionChange,
  placeholder = "Search for people...",
  disabled = false,
  maxSelectedOptions,
  width,
  className
}) => {
  const styles = useUserPickerStyles();
  const { searchUsers, getUsers } = useGraphAPIs({ context });
  const [availableUsers, setAvailableUsers] = useState<IUserProfile[]>([]);
  const [isSearching, setIsSearching] = useState(false);
  const [searchValue, setSearchValue] = useState("");
  
  // Infinite scroll state
  const [hasMore, setHasMore] = useState(true);
  const [isLoadingMore, setIsLoadingMore] = useState(false);
  const [nextPageToken, setNextPageToken] = useState<string | undefined>(undefined);
  const listRef = useRef<HTMLDivElement>(null);

  // Convert selected users to array of IDs for TagPicker
  const selectedUserIds = selectedUsers.map(user => user.id);

  /**
   * Load initial users when component mounts
   */
  useEffect(() => {
    const loadInitialUsers = async (): Promise<void> => {
      setIsSearching(true);
      try {
        const result = await getUsers(undefined, 20);
        setAvailableUsers(result.users);
        setHasMore(result.hasMore);
        setNextPageToken(result.nextPageToken);
      } catch (error) {
        console.error('Error loading initial users:', error);
        setAvailableUsers([]);
        setHasMore(false);
        setNextPageToken(undefined);
      } finally {
        setIsSearching(false);
      }
    };
    loadInitialUsers().catch(console.error);
  }, [getUsers]);

  /**
   * Handle search input changes
   */
  const handleInputChange = useCallback(async (value: string) => {
    setSearchValue(value);
    
    if (!value.trim()) {
      // If empty, show initial users
      try {
        const result = await getUsers(undefined, 20);
        setAvailableUsers(result.users);
        setHasMore(result.hasMore);
        setNextPageToken(result.nextPageToken);
      } catch (error) {
        console.error('Error loading users:', error);
      }
      return;
    }

    setIsSearching(true);
    try {
      const results = await searchUsers(value.trim(), 20);
      setAvailableUsers(results);
      // For search results, disable infinite scroll for now
      setHasMore(false);
      setNextPageToken(undefined);
    } catch (error) {
      console.error('Error searching users:', error);
      setAvailableUsers([]);
      setHasMore(false);
      setNextPageToken(undefined);
    } finally {
      setIsSearching(false);
    }
  }, [searchUsers, getUsers]);

  /**
   * Load more users for infinite scroll
   */
  const loadMoreUsers = useCallback(async () => {
    if (!hasMore || isLoadingMore || !nextPageToken || searchValue.trim()) {
      return;
    }

    setIsLoadingMore(true);
    try {
      const result = await getUsers(nextPageToken, 20);
      setAvailableUsers(prev => [...prev, ...result.users]);
      setHasMore(result.hasMore);
      setNextPageToken(result.nextPageToken);
    } catch (error) {
      console.error('Error loading more users:', error);
    } finally {
      setIsLoadingMore(false);
    }
  }, [getUsers, hasMore, isLoadingMore, nextPageToken, searchValue]);

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
      loadMoreUsers().catch(console.error);
    }
  }, [hasMore, isLoadingMore, searchValue, loadMoreUsers]);

  /**
   * Handle option selection and dismissal - follows the example pattern
   */
  const onOptionSelect: TagPickerProps["onOptionSelect"] = useCallback((e, data) => {
    if (data.value === "no-options" || data.value === "loading" || data.value === "loading-more") {
      return;
    }

    // Use the selectedOptions from data to update the selection
    const newSelectedUserIds = data.selectedOptions;
    const newSelectedUsers = selectedUsers.filter(user => 
      newSelectedUserIds.includes(user.id)
    );

    // If a new option was selected (not just dismissed), add it
    if (newSelectedUserIds.length > selectedUserIds.length) {
      const newUserId = newSelectedUserIds.find((id: string) => !selectedUserIds.includes(id));
      if (newUserId) {
        const selectedUser = availableUsers.find(user => user.id === newUserId);
        if (selectedUser) {
          // Check max limit
          if (maxSelectedOptions && selectedUsers.length >= maxSelectedOptions) {
            return;
          }
          newSelectedUsers.push(selectedUser);
        }
      }
      
      // Clear search input and reset to initial users after selection
      setSearchValue("");
      // Load initial users to reset the list
      getUsers(undefined, 20).then((result: IUsersPageResultSimplified) => {
        setAvailableUsers(result.users);
        setHasMore(result.hasMore);
        setNextPageToken(result.nextPageToken);
      }).catch(console.error);
    }

    onSelectionChange?.(newSelectedUsers);
  }, [availableUsers, selectedUsers, selectedUserIds, onSelectionChange, maxSelectedOptions, getUsers]);

  // Filter available users to exclude already selected ones
  const tagPickerOptions = availableUsers.filter(
    user => !selectedUserIds.includes(user.id)
  );

  return (
    <div 
      className={`${styles.container} ${className || ''}`}
      {...(width && { style: { width } })}
    >
      <div className={styles.tagPickerStyles}>
        <TagPicker
          onOptionSelect={onOptionSelect}
          selectedOptions={selectedUserIds}
        >
        <TagPickerControl>
          <TagPickerGroup aria-label="Selected Users">
            {selectedUsers.map((user) => (
              <Tag
                key={user.id}
                shape="circular"
                media={
                  <UserCard
                    userId={user.userPrincipalName}
                    avatarSize={20}
                    context={context}
                    avatarOnly={true}
                  />
                }
                value={user.id}
                dismissible
                disabled={disabled}
              >
                {user.displayName}
              </Tag>
            ))}
          </TagPickerGroup>
          <TagPickerInput 
            aria-label="Select Users"
            placeholder={placeholder}
            value={searchValue}
            onChange={(e) => handleInputChange(e.target.value)}
            disabled={disabled}
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
              {tagPickerOptions.map((user) => (
                <TagPickerOption
                  key={user.id}
                  value={user.id}
                  text={user.displayName}
                  media={
                    <UserCard
                      userId={user.userPrincipalName}
                      avatarSize={36}
                      context={context}
                      avatarOnly={true}
                      
                    />
                  }
                >
                  <div className={styles.optionContent}>
                    <Text weight="semibold">{user.displayName}</Text>
                    {user.jobTitle && (
                      <Text size={200} className={styles.optionSecondary}>
                        {user.jobTitle}
                      </Text>
                    )}
                    
                  </div>
                </TagPickerOption>
              ))}
              {isLoadingMore && !searchValue.trim() && (
                <TagPickerOption value="loading-more">
                  Loading more...
                </TagPickerOption>
              )}
            </>
          ) : (
            <TagPickerOption value="no-options">
              {searchValue ? "No users found" : "Start typing to search..."}
            </TagPickerOption>
          )}
        </TagPickerList>
        </TagPicker>
      </div>
    </div>
  );
};
