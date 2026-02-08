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
import { useCallback, useState } from "react";
import { ETargetTypes } from "../../constants";
import { useResourcePickerStyles } from "./useResourcePickerStyles";
import {
  Person24Regular,
  PeopleTeam24Regular,
  Phone24Regular,
  Calendar24Regular,
  Mail24Regular,
  Organization24Regular,
  ContactCard24Regular,
  Shield24Regular,
  Chat24Regular,
  TaskListSquareLtr24Regular,
  CheckboxChecked24Regular,
} from "@fluentui/react-icons";

export interface ITargetTypePickerProps {
  availableTargetTypes: string[];
  selectedTargetType?: string;
  onSelectionChange?: (selectedTargetType: string | undefined) => void;
  placeholder?: string;
  disabled?: boolean;
  width?: string | number;
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
    case ETargetTypes.Post:
      return <Chat24Regular />;
    case ETargetTypes.TodoTask:
      return <CheckboxChecked24Regular />;
    case ETargetTypes.TodoTaskList:
      return <TaskListSquareLtr24Regular />;
    default:
      return <Person24Regular />;
  }
};

/**
 * Get description for target type
 */
const getTargetTypeDescription = (targetType: string): string => {
  switch (targetType) {
    case ETargetTypes.User:
      return "Microsoft 365 users in your organization";
    case ETargetTypes.Group:
      return "Microsoft 365 groups and security groups";
    case ETargetTypes.Device:
      return "Devices registered in Microsoft Entra ID";
    case ETargetTypes.Event:
      return "Calendar events for users and groups";
    case ETargetTypes.Message:
      return "Email messages in mailboxes";
    case ETargetTypes.Organization:
      return "Organization settings and configuration";
    case ETargetTypes.Contact:
      return "Organizational contacts";
    case ETargetTypes.AdministrativeUnit:
      return "Administrative units in Microsoft Entra ID";
    case ETargetTypes.Post:
      return "Group conversation posts";
    case ETargetTypes.TodoTask:
      return "To Do tasks";
    case ETargetTypes.TodoTaskList:
      return "To Do task lists";
    default:
      return "";
  }
};

export const TargetTypePicker: React.FunctionComponent<ITargetTypePickerProps> = ({
  availableTargetTypes,
  selectedTargetType,
  onSelectionChange,
  placeholder = "Select target type...",
  disabled = false,
  width,
}) => {
  const styles = useResourcePickerStyles();
  const [searchValue, setSearchValue] = useState("");

  // Filter target types based on search
  const filteredTargetTypes = React.useMemo(() => {
    if (!searchValue.trim()) {
      return availableTargetTypes;
    }
    const query = searchValue.toLowerCase();
    return availableTargetTypes.filter(
      (type) =>
        type.toLowerCase().includes(query) ||
        getTargetTypeDescription(type).toLowerCase().includes(query)
    );
  }, [availableTargetTypes, searchValue]);

  // Get options that are not already selected
  const tagPickerOptions = React.useMemo(() => {
    return filteredTargetTypes.filter((type) => type !== selectedTargetType);
  }, [filteredTargetTypes, selectedTargetType]);

  // Handle selection change
  const handleOptionSelect: TagPickerProps["onOptionSelect"] = useCallback(
    (_event, data) => {
      if (data.selectedOptions.length === 0) {
        onSelectionChange?.(undefined);
      } else {
        // Get the last selected option (since we allow only one)
        const selectedValue = data.selectedOptions[data.selectedOptions.length - 1];
        onSelectionChange?.(selectedValue);
      }
      setSearchValue("");
    },
    [onSelectionChange]
  );

  // Get selected options array for TagPicker
  const selectedOptions = selectedTargetType ? [selectedTargetType] : [];

  return (
    <div
      className={styles.container}
      {...(width && { style: { width } })}
    >
      <TagPicker
          onOptionSelect={handleOptionSelect}
          selectedOptions={selectedOptions}
          disabled={disabled}
        >
          <TagPickerControl className={styles.tagPickerControl}>
            <TagPickerGroup aria-label="Selected Target Type">
              {selectedTargetType && (
                <Tag
                  key={selectedTargetType}
                  shape="circular"
                  media={
                    <Avatar
                      name={selectedTargetType}
                      size={20}
                      icon={getTargetTypeIcon(selectedTargetType)}
                      color="colorful"
                    />
                  }
                  value={selectedTargetType}
                  dismissible
                  disabled={disabled}
                >
                  {selectedTargetType}
                </Tag>
              )}
            </TagPickerGroup>
            <TagPickerInput
              className={styles.tagPickerInput}
              aria-label="Select Target Type"
              placeholder={selectedTargetType ? "" : placeholder}
              value={searchValue}
              onChange={(e) => setSearchValue(e.target.value)}
              disabled={disabled}
            />
          </TagPickerControl>

          <TagPickerList className={styles.pickerList}>
            {tagPickerOptions.length > 0 ? (
              <>
                {tagPickerOptions.map((targetType) => {
                  const description = getTargetTypeDescription(targetType);

                  return (
                    <TagPickerOption
                      key={targetType}
                      value={targetType}
                      text={targetType}
                      media={
                        <Avatar
                          name={targetType}
                          size={36}
                          icon={getTargetTypeIcon(targetType)}
                          color="colorful"
                        />
                      }
                    >
                      <div className={styles.optionContent}>
                        <Text weight="semibold">{targetType}</Text>
                        {description && (
                          <Text size={200} className={styles.optionSecondary}>
                            {description}
                          </Text>
                        )}
                      </div>
                    </TagPickerOption>
                  );
                })}
              </>
            ) : (
              <TagPickerOption value="no-options" text="No target types available">
                {searchValue
                  ? "No matching target types found"
                  : "No target types available"}
              </TagPickerOption>
            )}
          </TagPickerList>
        </TagPicker>
    </div>
  );
};

export default TargetTypePicker;
