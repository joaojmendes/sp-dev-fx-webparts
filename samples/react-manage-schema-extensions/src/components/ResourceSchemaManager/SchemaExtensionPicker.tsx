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
  Badge,
} from "@fluentui/react-components";
import { useCallback, useEffect, useState } from "react";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { ISchemaExtension } from "../../models/ISchemaExtension";
import { useResourceSchemaExtension } from "../../hooks/useResourceSchemaExtension";
import { useResourcePickerStyles } from "./useResourcePickerStyles";
import { DataUsageSettings24Regular } from "@fluentui/react-icons";
import {   useAtomValue } from "jotai";
import { appGlobalStateAtom } from "../../atoms/appGlobalState";

export interface ISchemaExtensionPickerProps {
 
  selectedSchemaExtension?: ISchemaExtension;
  onSelectionChange?: (schemaExtension: ISchemaExtension | undefined) => void;
  placeholder?: string;
  disabled?: boolean;
  width?: string | number;
  className?: string;
}

export const SchemaExtensionPicker: React.FunctionComponent<ISchemaExtensionPickerProps> = ({
  
  selectedSchemaExtension,
  onSelectionChange,
  placeholder = "Search for schema extensions...",
  disabled = false,
  width,
  className,
}) => {
  const styles = useResourcePickerStyles();
  const appGlobalState = useAtomValue(appGlobalStateAtom);
  
  const { context } = appGlobalState || {};
  const { getSchemaExtensions } = useResourceSchemaExtension({ context: context as BaseComponentContext });

  const [availableSchemas, setAvailableSchemas] = useState<ISchemaExtension[]>([]);
  const [filteredSchemas, setFilteredSchemas] = useState<ISchemaExtension[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [searchValue, setSearchValue] = useState("");
  const [hasLoaded, setHasLoaded] = useState(false);

  // Convert selected schema to array of IDs for TagPicker
  const selectedSchemaIds = selectedSchemaExtension?.id ? [selectedSchemaExtension.id] : [];

  /**
   * Load schema extensions when component mounts
   */
  useEffect(() => {
    let isCancelled = false;

    const loadSchemas = async (): Promise<void> => {
      setIsLoading(true);
      try {
        const schemas = await getSchemaExtensions();
        if (!isCancelled) {
          setAvailableSchemas(schemas);
          setFilteredSchemas(schemas);
          setHasLoaded(true);
        }
      } catch (error) {
        console.error("Error loading schema extensions:", error);
        if (!isCancelled) {
          setAvailableSchemas([]);
          setFilteredSchemas([]);
          setHasLoaded(true);
        }
      } finally {
        if (!isCancelled) {
          setIsLoading(false);
        }
      }
    };

    loadSchemas().catch(console.error);

    return () => {
      isCancelled = true;
    };
  }, [getSchemaExtensions]);

  /**
   * Handle search input changes - filter locally
   */
  const handleInputChange = useCallback(
    (value: string) => {
      setSearchValue(value);

      if (!value.trim()) {
        setFilteredSchemas(availableSchemas);
        return;
      }

      const searchLower = value.toLowerCase().trim();
      const filtered = availableSchemas.filter(
        (schema) =>
          schema.id?.toLowerCase().includes(searchLower) ||
          schema.description?.toLowerCase().includes(searchLower)
      );
      setFilteredSchemas(filtered);
    },
    [availableSchemas]
  );

  /**
   * Handle option selection and dismissal
   */
  const onOptionSelect: TagPickerProps["onOptionSelect"] = useCallback(
    (e, data) => {
      if (data.value === "no-options" || data.value === "loading") {
        return;
      }

      const newSelectedIds = data.selectedOptions;

      // If selection is empty, clear the selection
      if (newSelectedIds.length === 0) {
        onSelectionChange?.(undefined);
        return;
      }

      // Find the newly selected schema (we only allow single selection)
      const selectedId = newSelectedIds[0];
      const selected = availableSchemas.find((schema) => schema.id === selectedId);

      if (selected) {
        onSelectionChange?.(selected);
        // Clear search input after selection
        setSearchValue("");
        setFilteredSchemas(availableSchemas);
      }
    },
    [availableSchemas, onSelectionChange]
  );

  // Filter available schemas to exclude already selected one
  const tagPickerOptions = filteredSchemas.filter(
    (schema) => schema.id && !selectedSchemaIds.includes(schema.id)
  );

  /**
   * Get status badge color
   */
  const getStatusColor = (status: string): "success" | "warning" | "informative" => {
    switch (status) {
      case "Available":
        return "success";
      case "InDevelopment":
        return "warning";
      default:
        return "informative";
    }
  };

  return (
    <div
      className={`${styles.container} ${className || ""}`}
      {...(width && { style: { width } })}
    >
      <TagPicker
          onOptionSelect={onOptionSelect}
          selectedOptions={selectedSchemaIds}
          disabled={disabled}
        >
          <TagPickerControl className={styles.tagPickerControl}>
            <TagPickerGroup aria-label="Selected Schema Extension">
              {selectedSchemaExtension && (
                <Tag
                  key={selectedSchemaExtension.id}
                  shape="circular"
                  media={
                    <Avatar
                      icon={<DataUsageSettings24Regular />}
                      size={20}
                      color="brand"
                    />
                  }
                  value={selectedSchemaExtension.id}
                  dismissible
                  disabled={disabled}
                >
                  {selectedSchemaExtension.id}
                </Tag>
              )}
            </TagPickerGroup>
            <TagPickerInput
              className={styles.tagPickerInput}
              aria-label="Select Schema Extension"
              placeholder={selectedSchemaExtension ? "" : placeholder}
              value={searchValue}
              onChange={(e) => handleInputChange(e.target.value)}
              disabled={disabled}
            />
          </TagPickerControl>

          <TagPickerList className={styles.pickerList}>
            {isLoading ? (
              <TagPickerOption value="loading">Loading schema extensions...</TagPickerOption>
            ) : tagPickerOptions.length > 0 ? (
              <>
                {tagPickerOptions.map((schema) => (
                  <TagPickerOption
                    key={schema.id}
                    value={schema.id || ""}
                    text={schema.id || ""}
                    media={
                      <Avatar
                        icon={<DataUsageSettings24Regular />}
                        size={36}
                        color="brand"
                      />
                    }
                  >
                    <div className={styles.optionContent}>
                      <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                        <Text weight="semibold">{schema.id}</Text>
                        <Badge
                          appearance="filled"
                          color={getStatusColor(schema.status || "")}
                          size="small"
                        >
                          {schema.status}
                        </Badge>
                      </div>
                      <Text size={200} className={styles.optionSecondary}>
                        {schema.description || "No description"} •{" "}
                        {schema.properties?.length || 0} properties •{" "}
                        {schema.targetTypes?.join(", ") || "No target types"}
                      </Text>
                    </div>
                  </TagPickerOption>
                ))}
              </>
            ) : (
              <TagPickerOption value="no-options">
                {!hasLoaded
                  ? "Loading..."
                  : searchValue
                  ? "No schema extensions found"
                  : availableSchemas.length === 0
                  ? "No schema extensions available"
                  : "Start typing to search..."}
              </TagPickerOption>
            )}
          </TagPickerList>
        </TagPicker>
    </div>
  );
};
