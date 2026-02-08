/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from "react";
import {
  Field,
  Input,
  Switch,
  SpinButton,
  Spinner,
  InfoLabel,
  tokens,
} from "@fluentui/react-components";
import {
  Info24Regular,
  Calendar20Regular,
  ToggleLeft20Regular,
  NumberSymbol20Regular,
  TextT20Regular,
  Document20Regular,
} from "@fluentui/react-icons";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import {
  ISchemaExtension,
  ISchemaExtensionProperty,
} from "../../models/ISchemaExtension";
import {
  IGraphResource,
  IResourceAttributeData,
  useResourceSchemaExtension,
} from "../../hooks/useResourceSchemaExtension";
import { EPropertyTypes } from "../../constants";
import {
  StackV2 as Stack,
  RenderLabel,
  EMessageType,
  ShowMessage,
  TypographyControl,
} from "@spteck/react-controls";
import { useResourceAttributeManagerStyles } from "./useResourceAttributeManagerStyles";

export interface IResourceAttributeManagerProps {
  context: BaseComponentContext;
  schemaExtension: ISchemaExtension;
  resource: IGraphResource;
  targetType: string;
  onSuccess?: (resourceId: string, data: IResourceAttributeData) => void;
  onCancel?: () => void;
  hideButtons?: boolean;
  onSaveRef?: React.MutableRefObject<(() => void) | undefined>;
  onResetRef?: React.MutableRefObject<(() => void) | undefined>;
  onStateChange?: (state: {
    isSaving: boolean;
    isLoading: boolean;
    hasChanges: boolean;
  }) => void;
  className?: string;
  hasMultipleTargetTypes?: boolean;
}

// Get icon for property type
const getPropertyTypeIcon = (type: string): React.ReactElement => {
  switch (type) {
    case EPropertyTypes.Boolean:
      return <ToggleLeft20Regular />;
    case EPropertyTypes.Integer:
      return <NumberSymbol20Regular />;
    case EPropertyTypes.DateTime:
      return <Calendar20Regular />;
    case EPropertyTypes.Binary:
      return <Document20Regular />;
    case EPropertyTypes.String:
    default:
      return <TextT20Regular />;
  }
};

/**
 * Get default value based on property type
 */
const getDefaultValue = (type: string): string | number | boolean => {
  switch (type) {
    case EPropertyTypes.Boolean:
      return false;
    case EPropertyTypes.Integer:
      return 0;
    case EPropertyTypes.DateTime:
      return "";
    case EPropertyTypes.Binary:
      return "";
    case EPropertyTypes.String:
    default:
      return "";
  }
};

/**
 * Parse value based on property type
 */
const parseValue = (
  value: unknown,
  type: string,
): string | number | boolean => {
  if (value === null || value === undefined) {
    return getDefaultValue(type);
  }

  switch (type) {
    case EPropertyTypes.Boolean:
      return Boolean(value);
    case EPropertyTypes.Integer:
      return typeof value === "number"
        ? value
        : parseInt(String(value), 10) || 0;
    case EPropertyTypes.DateTime:
      return String(value);
    case EPropertyTypes.Binary:
      return String(value);
    case EPropertyTypes.String:
    default:
      return String(value);
  }
};

export const ResourceAttributeManager: React.FunctionComponent<
  IResourceAttributeManagerProps
> = ({
  context,
  schemaExtension,
  resource,
  targetType,
  onSuccess,
  onCancel,
  hideButtons,
  onSaveRef,
  onResetRef,
  onStateChange,
  hasMultipleTargetTypes = false,
  className,
}) => {
  const { getResourceAttributes, updateResourceAttributes, error } =
    useResourceSchemaExtension({ context });
  const styles = useResourceAttributeManagerStyles(hasMultipleTargetTypes);

  // State
  const [formData, setFormData] = React.useState<IResourceAttributeData>({});
  const [originalData, setOriginalData] =
    React.useState<IResourceAttributeData>({});
  const [isLoadingData, setIsLoadingData] = React.useState<boolean>(true);
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [successMessage, setSuccessMessage] = React.useState<string>("");
  const [errorMessage, setErrorMessage] = React.useState<string>("");

  // Get properties from schema extension
  const properties = schemaExtension?.properties || [];

  // Load current attribute values when resource or schema changes
  React.useEffect(() => {
    let isCancelled = false;

    const loadAttributeData = async (): Promise<void> => {
      if (!resource?.id || !schemaExtension?.id) {
        setIsLoadingData(false);
        return;
      }

      setIsLoadingData(true);
      setSuccessMessage("");
      setErrorMessage("");

      try {
        const data = await getResourceAttributes(
          targetType,
          resource.id,
          schemaExtension.id,
        );

        if (!isCancelled) {
          // Initialize form data with loaded values or defaults
          const initialData: IResourceAttributeData = {};
          properties.forEach((prop) => {
            initialData[prop.name] = parseValue(data[prop.name], prop.type);
          });

          setFormData(initialData);
          setOriginalData(initialData);
        }
      } catch (err) {
        console.error("Error loading attribute data:", err);
        if (!isCancelled) {
          // Initialize with default values on error
          const defaultData: IResourceAttributeData = {};
          properties.forEach((prop) => {
            defaultData[prop.name] = getDefaultValue(prop.type);
          });
          setFormData(defaultData);
          setOriginalData(defaultData);
        }
      } finally {
        if (!isCancelled) {
          setIsLoadingData(false);
        }
      }
    };

    loadAttributeData().catch(console.error);

    return () => {
      isCancelled = true;
    };
  }, [
    resource?.id,
    schemaExtension?.id,
    targetType,
    getResourceAttributes,
    properties,
  ]);

  // Handle input changes based on property type
  const handleInputChange = React.useCallback(
    (propertyName: string, value: string | number | boolean, type: string) => {
      setFormData((prev) => ({
        ...prev,
        [propertyName]: parseValue(value, type),
      }));
      setSuccessMessage("");
    },
    [],
  );

  // Handle form submission
  const handleSubmit = React.useCallback(async () => {
    if (!resource?.id || !schemaExtension?.id) {
      setErrorMessage("Resource or schema extension not selected");
      return;
    }

    setIsSaving(true);
    setSuccessMessage("");
    setErrorMessage("");

    try {
      await updateResourceAttributes(
        targetType,
        resource.id,
        schemaExtension.id,
        formData,
      );
      setSuccessMessage(
        `Successfully updated attributes for ${resource.displayName || resource.id}`,
      );
      setOriginalData({ ...formData });
      onSuccess?.(resource.id, formData);
    } catch (err) {
      console.error("Error updating attributes:", err);
      setErrorMessage(
        err instanceof Error ? err.message : "Failed to update attributes",
      );
    } finally {
      setIsSaving(false);
    }
  }, [
    resource,
    schemaExtension,
    targetType,
    formData,
    updateResourceAttributes,
    onSuccess,
  ]);

  // Handle reset
  const handleReset = React.useCallback(() => {
    setFormData({ ...originalData });
    setSuccessMessage("");
    setErrorMessage("");
  }, [originalData]);

  // Check if form has changes
  const hasChanges = React.useMemo(() => {
    return JSON.stringify(formData) !== JSON.stringify(originalData);
  }, [formData, originalData]);

  // Expose handlers via refs for external access
  React.useEffect(() => {
    if (onSaveRef) {
      onSaveRef.current = () => {
        // call handleSubmit
        handleSubmit();
      };
    }
    if (onResetRef) {
      onResetRef.current = handleReset;
    }
  }, [handleSubmit, handleReset]);

  // Notify parent of state changes
  React.useEffect(() => {
    onStateChange?.({ isSaving, isLoading: isLoadingData, hasChanges });
  }, [isSaving, isLoadingData, hasChanges, onStateChange]);

  // Render input based on property type
  const renderPropertyInput = (
    property: ISchemaExtensionProperty,
  ): JSX.Element => {
    const value = formData[property.name];

    switch (property.type) {
      case EPropertyTypes.Boolean:
        return (
          <Field>
            <Switch
              checked={Boolean(value)}
              onChange={(_, data) =>
                handleInputChange(property.name, data.checked, property.type)
              }
              disabled={isSaving}
            />
          </Field>
        );

      case EPropertyTypes.Integer:
        return (
          <Field>
            <SpinButton
              value={typeof value === "number" ? value : 0}
              onChange={(_, data) =>
                handleInputChange(property.name, data.value || 0, property.type)
              }
              disabled={isSaving}
            />
          </Field>
        );

      case EPropertyTypes.DateTime:
        return (
          <Field hint="Format: YYYY-MM-DDTHH:mm:ssZ">
            <Input
              type="datetime-local"
              value={String(value || "")
                .replace("Z", "")
                .slice(0, 16)}
              onChange={(_, data) => {
                const isoValue = data.value ? `${data.value}:00Z` : "";
                handleInputChange(property.name, isoValue, property.type);
              }}
              disabled={isSaving}
            />
          </Field>
        );

      case EPropertyTypes.Binary:
        return (
          <Field hint="Base64 encoded binary data">
            <Input
              value={String(value || "")}
              onChange={(_, data) =>
                handleInputChange(property.name, data.value, property.type)
              }
              disabled={isSaving}
              placeholder="Enter base64 encoded data..."
            />
          </Field>
        );

      case EPropertyTypes.String:
      default:
        return (
          <Field>
            <Input
              value={String(value || "")}
              onChange={(_, data) =>
                handleInputChange(property.name, data.value, property.type)
              }
              disabled={isSaving}
              placeholder={`Enter ${property.name}...`}
            />
          </Field>
        );
    }
  };

  // Loading state
  if (isLoadingData) {
    return (
      <Stack gap="10px" padding="m">
        <Stack
          alignItems="center"
          justifyContent="center"
          padding="xl"
          gap="10px"
        >
          <Spinner size="medium" label="Loading attribute data..." />
        </Stack>
      </Stack>
    );
  }

  // No properties defined
  if (properties.length === 0) {
    return (
      <Stack gap="10px" padding="m">
        <Stack
          alignItems="center"
          justifyContent="center"
          padding="xl"
          gap="10px"
        >
          <Info24Regular />
          <TypographyControl fontSize="400" fontWeight="semibold">
            No properties defined
          </TypographyControl>
          <TypographyControl fontSize="200" fontWeight="regular" color="subtle">
            This schema extension has no properties configured.
          </TypographyControl>
        </Stack>
      </Stack>
    );
  }

  return (
    <Stack gap="10px" padding="s" className={styles.root}>
        {/* Success Message */}
      {successMessage && (
        <ShowMessage
          message={successMessage}
          messageType={EMessageType.SUCCESS}
        />
      )}

      {/* Error Message */}
      {(errorMessage || error) && (
        <ShowMessage
          message={errorMessage || error || ""}
          messageType={EMessageType.ERROR}
        />
      )}
      {/* Attributes  */}
      <Stack gap="10px" padding="s" className={styles.innerContent}>
        <Stack gap="5px" direction="horizontal">
          <TypographyControl
            fontSize="400"
            fontWeight="semibold"
            color={tokens.colorBrandBackground}
          >
            Schema Properties ({properties.length})
          </TypographyControl>
        </Stack>
        <div className={styles.propertiesContainer}>
        {properties.map((property) => (
          <Stack key={property.name} gap="5px" padding="s">
            <Stack gap="5px" direction="horizontal">
              <RenderLabel
                label={property.name}
                icon={getPropertyTypeIcon(property.type)}
              />
              <InfoLabel info={`Type: ${property.type}`} />
            </Stack>
            {renderPropertyInput(property)}
          </Stack>
        ))}
        </div>
      </Stack>
      
    
    </Stack>
  );
};
