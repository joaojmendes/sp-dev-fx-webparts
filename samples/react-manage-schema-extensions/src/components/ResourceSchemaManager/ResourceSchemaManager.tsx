import * as React from "react";
import {
  Field,
  Text,
  Card,
  CardHeader,
  Badge,
  Divider,
  InlineDrawer,
  DrawerHeader,
  DrawerHeaderTitle,
  DrawerBody,
  DrawerFooter,
  Button,
  Body1,
  Subtitle1,
} from "@fluentui/react-components";
import {
  Dismiss24Regular,
  DataBarHorizontal20Regular,
} from "@fluentui/react-icons";
 
import { ISchemaExtension } from "../../models/ISchemaExtension";
import {
  IGraphResource,
  IResourceAttributeData,
} from "../../hooks/useResourceSchemaExtension";
import { ResourcePicker } from "./ResourcePicker";
import { ResourceAttributeManager } from "./ResourceAttributeManager";
import { SchemaExtensionPicker } from "./SchemaExtensionPicker";
import { TargetTypePicker } from "./TargetTypePicker";
import { useResourceSchemaManagerStyles } from "./useResourceSchemaManagerStyles";
import { StackV2, TypographyControl } from "@spteck/react-controls";
import { UserPicker } from "../UserPicker/UserPicker";
import { IUserProfile } from "../../models/IUserData";
import { ETargetTypes } from "../../constants";
import { useAppToast } from "@spteck/m365-hooks";
import { useAtomValue } from "jotai";
import { appGlobalStateAtom } from "../../atoms/appGlobalState";
import { BaseComponentContext } from "@microsoft/sp-component-base";

export interface IResourceSchemaManagerProps {
  isOpen: boolean;
  onClose: () => void;
  onSuccess?: (resourceId: string, data: IResourceAttributeData) => void;
}

export const ResourceSchemaManager: React.FunctionComponent<
  IResourceSchemaManagerProps
> = ({ isOpen, onClose, onSuccess }) => {
  const styles = useResourceSchemaManagerStyles();
  const { showSuccessToast } = useAppToast();
  const appGlobalState = useAtomValue(appGlobalStateAtom);
  const { context } = appGlobalState || {};
  // State
  const [selectedSchema, setSelectedSchema] = React.useState<
    ISchemaExtension | undefined
  >(undefined);
  const [selectedTargetType, setSelectedTargetType] = React.useState<
    string | undefined
  >(undefined);
  const [selectedResources, setSelectedResources] = React.useState<
    IGraphResource[]
  >([]);
  const [saveSuccessful, setSaveSuccessful] = React.useState<boolean>(false);

  // State for footer buttons
  const [attributeState, setAttributeState] = React.useState<{
    isSaving: boolean;
    isLoading: boolean;
    hasChanges: boolean;
  }>({ isSaving: false, isLoading: false, hasChanges: false });

  // Refs for action handlers
  const saveRef = React.useRef<(() => void) | undefined>();
  const resetRef = React.useRef<(() => void) | undefined>();

  // Reset form when drawer opens
  React.useEffect(() => {
    if (isOpen) {
      setSelectedSchema(undefined);
      setSelectedTargetType(undefined);
      setSelectedResources([]);
      setSaveSuccessful(false);
      setAttributeState({
        isSaving: false,
        isLoading: false,
        hasChanges: false,
      });
    }
  }, [isOpen]);

  // Check if schema has multiple target types
  const hasMultipleTargetTypes = React.useMemo(() => {
    return (selectedSchema?.targetTypes?.length || 0) > 1;
  }, [selectedSchema]);

  // Auto-select target type if schema has only one
  React.useEffect(() => {
    if (selectedSchema?.targetTypes?.length === 1) {
      setSelectedTargetType(selectedSchema.targetTypes[0]);
    } else if (!selectedSchema) {
      setSelectedTargetType(undefined);
    }
  }, [selectedSchema]);

  // Handle schema selection from picker
  const handleSchemaSelectionChange = React.useCallback(
    (schema: ISchemaExtension | undefined) => {
      setSelectedSchema(schema);
      setSelectedTargetType(undefined);
      setSelectedResources([]);
    },
    [],
  );

  // Handle target type selection
  const handleTargetTypeChange = React.useCallback(
    (targetType: string | undefined) => {
      setSelectedTargetType(targetType);
      setSelectedResources([]);
      setSaveSuccessful(false);
    },
    [],
  );

  // Handle resource selection
  const handleResourceSelection = React.useCallback(
    (resources: IGraphResource[]) => {
      setSelectedResources(resources);
    },
    [],
  );

  // Handle success callback - close drawer after save
  const handleSuccess = React.useCallback(
    (resourceId: string, data: IResourceAttributeData) => {
      setSaveSuccessful(true);
      showSuccessToast("Attributes saved successfully");
      onSuccess?.(resourceId, data);
      // Close the drawer after a short delay to show success state
      setTimeout(() => {
        onClose();
      }, 1000);
    },
    [onSuccess, onClose],
  );

  // Handle schema selection - reset save state
  const handleSchemaChange = React.useCallback(
    (schema: ISchemaExtension | undefined) => {
      handleSchemaSelectionChange(schema);
      setSaveSuccessful(false);
    },
    [handleSchemaSelectionChange],
  );

  // Handle resource selection - reset save state
  const handleResourceChange = React.useCallback(
    (resources: IGraphResource[]) => {
      handleResourceSelection(resources);
      setSaveSuccessful(false);
    },
    [handleResourceSelection],
  );

  // Get selected resource (first one, since we typically select one)
  const selectedResource =
    selectedResources.length > 0 ? selectedResources[0] : undefined;

  const RenderTitle = React.useCallback(() => {
    return (
      <StackV2 direction="horizontal" gap="m">
        <DataBarHorizontal20Regular style={{ width: "42px", height: "42px" }} />
        <StackV2>
          <Subtitle1>Manage Resource Attributes</Subtitle1>
          <Body1>Add attributes to the selected resource</Body1>
        </StackV2>
      </StackV2>
    );
  }, []);

  return (
    <InlineDrawer
      className={styles.drawerRoot}
      open={isOpen}
      position="end"
      style={{ width: "550px" }}
    >
      <DrawerHeader>
        <DrawerHeaderTitle
          action={
            <Button
              appearance="subtle"
              aria-label="Close"
              icon={<Dismiss24Regular />}
              onClick={() => onClose()}
            />
          }
        >
          <RenderTitle />
        </DrawerHeaderTitle>
      </DrawerHeader>
      <Divider style={{ flexGrow: 0 }} />
      <DrawerBody className={styles.drawerContent}>
        <StackV2 className={styles.drawerBodyInner}>
          {/*Select Schema Extension */}
          <StackV2 className={styles.section}>
            <Text className={styles.sectionTitle}>Select Schema Extension</Text>

            <Field label="Schema Extension" required>
              <SchemaExtensionPicker
                 
                selectedSchemaExtension={selectedSchema}
                onSelectionChange={handleSchemaChange}
                placeholder="Search for schema extensions..."
              />
            </Field>

            {/* Schema Info */}
            <StackV2>
              {selectedSchema && (
                <Card size="small" className={styles.infoCard}>
                  <CardHeader
                    header={<Text weight="semibold">Schema Details</Text>}
                    description={
                      <TypographyControl fontSize="300">
                        {selectedSchema.description ||
                          "No description provided"}
                      </TypographyControl>
                    }
                  />

                  <StackV2
                    direction="horizontal"
                    gap="m"
                    alignItems="center"
                    justifyContent="start"
                  >
                    <Text size={200}>Status: </Text>
                    <Badge
                      appearance="filled"
                      color={
                        selectedSchema.status === "Available"
                          ? "success"
                          : "warning"
                      }
                    >
                      {selectedSchema.status}
                    </Badge>
                    <Text size={200}> • Target Types: </Text>
                    {selectedSchema.targetTypes?.map((targetType) => (
                      <Badge
                        key={targetType}
                        appearance="filled"
                        color="brand"
                        size="medium"
                      >
                        {targetType}
                      </Badge>
                    ))}
                  </StackV2>
                </Card>
              )}
            </StackV2>
          </StackV2>

          {/* Select Target Type - Only show when schema has multiple target types */}
          {selectedSchema && hasMultipleTargetTypes && (
            <StackV2 className={styles.section}>
              <Text className={styles.sectionTitle}>Select Target Type</Text>
              <Field label="Target Type" required>
                <TargetTypePicker
                  availableTargetTypes={selectedSchema.targetTypes || []}
                  selectedTargetType={selectedTargetType}
                  onSelectionChange={handleTargetTypeChange}
                  placeholder="Select a target type..."
                />
              </Field>
            </StackV2>
          )}

          {/* Select Resource */}
          {selectedSchema && selectedTargetType && (
            <StackV2 className={styles.section}>
              <Text className={styles.sectionTitle}>
                Select {selectedTargetType}
              </Text>

              <Field label={`${selectedTargetType}`} required>
                {selectedTargetType === ETargetTypes.User ? (
                  <UserPicker
                    context={context as BaseComponentContext}
                    selectedUsers={selectedResources.map(
                      (r) =>
                        ({
                          id: r.id,
                          displayName: r.displayName || "",
                          mail: (r.mail as string) || "",
                          userPrincipalName:
                            (r.userPrincipalName as string) || "",
                          jobTitle: (r.jobTitle as string) || "",
                        }) as IUserProfile,
                    )}
                    onSelectionChange={(users) => {
                      const resources: IGraphResource[] = users.map((user) => ({
                        id: user.id,
                        displayName: user.displayName,
                        mail: user.mail,
                        userPrincipalName: user.userPrincipalName,
                        jobTitle: user.jobTitle,
                      }));
                      handleResourceChange(resources);
                    }}
                    maxSelectedOptions={1}
                    placeholder={`Search for users...`}
                  />
                ) : (
                  <ResourcePicker
                    context={context as BaseComponentContext}
                    targetType={selectedTargetType}
                    selectedResources={selectedResources}
                    onSelectionChange={handleResourceChange}
                    maxSelectedOptions={1}
                    placeholder={`Search for ${selectedTargetType.toLowerCase()}s...`}
                  />
                )}
              </Field>
            </StackV2>
          )}

          {/* Manage Attributes */}
          {selectedSchema && selectedTargetType && selectedResource && (
            <>
              <Divider className={styles.divider} />
              <StackV2 className={styles.attributeManagerSection}>
                <Text className={styles.sectionTitle}> Manage Attributes</Text>

                <ResourceAttributeManager
                  context={context as BaseComponentContext}
                  schemaExtension={selectedSchema}
                  resource={selectedResource}
                  targetType={selectedTargetType}
                  onSuccess={handleSuccess}
                  hideButtons={true}
                  onSaveRef={saveRef}
                  onResetRef={resetRef}
                  onStateChange={setAttributeState}
                  hasMultipleTargetTypes={hasMultipleTargetTypes}
                />
              </StackV2>
            </>
          )}
        </StackV2>
      </DrawerBody>
      <DrawerFooter className={styles.drawerFooter}>
        <Button appearance="secondary" onClick={onClose}>
          Close
        </Button>
        {selectedResource && (
          <>
            <Button
              appearance="secondary"
              onClick={() => resetRef.current?.()}
              disabled={
                attributeState.isSaving ||
                !attributeState.hasChanges ||
                saveSuccessful
              }
            >
              Reset
            </Button>
            <Button
              appearance="primary"
              onClick={() => saveRef.current?.()}
              disabled={
                attributeState.isSaving ||
                attributeState.isLoading ||
                !attributeState.hasChanges ||
                saveSuccessful
              }
            >
              {attributeState.isSaving
                ? "Saving..."
                : saveSuccessful
                  ? "Saved!"
                  : "Save Attributes"}
            </Button>
          </>
        )}
      </DrawerFooter>
    </InlineDrawer>
  );
};
