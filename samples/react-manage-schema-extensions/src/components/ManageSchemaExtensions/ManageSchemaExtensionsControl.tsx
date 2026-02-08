import * as React from "react";

import { Divider, Subtitle1 } from "@fluentui/react-components";
import {
  Add20Regular,
  ArrowClockwise20Regular,
  ArrowSync20Regular,
  DatabaseSearch20Regular,
  Delete20Regular,
  Edit20Regular,
  Eye20Regular,
} from "@fluentui/react-icons";

import { ChangeSchemaStatus } from "../ChangeSchemaStatus";
import { DeleteSchemaExtension } from "../DeleteSchemaExtension";
import { IManageSchemaExtensionsProps } from "./IManageSchemaExtensionsProps";
import { ISchemaExtension } from "../../models/ISchemaExtension";
import { InformationPanel } from "./InformationPanel";
import SchemaExtensionDrawer from "../SchemaExtensionDrawer/SchemaExtensionDrawer";
import { SchemaExtensionViewer } from "../SchemaExtensionViewer";
import SchemaExtensionsListView from "../SchemaExtensionsListView/SchemaExtensionsListView";
import { ListToolbar, IToolbarItem } from "../ListToolbar";
import { SchemaStatusRestrictionDialog } from "../SchemaExtensionDrawer";
import { ResourceSchemaManager } from "../ResourceSchemaManager";
import { StackV2 as Stack } from "@spteck/react-controls";
import { appGlobalStateAtom } from "../../atoms/appGlobalState";
import { useAtom } from "jotai";
import { useLogging } from "@spteck/m365-hooks";
import { useManageSchemaExtensionsStyles } from "./useManageSchemaExtensionsStyles";
import { useAppToast } from "@spteck/m365-hooks";

export enum EActiveDrawer {
  Create = "create",
  Edit = "edit",
  View = "view",
  ResourceManager = "resourceManager",
}

export const ManageSchemaExtensionsControl: React.FunctionComponent<
  IManageSchemaExtensionsProps
> = (props) => {
  const [appGlobalState] = useAtom(appGlobalStateAtom);
  const styles = useManageSchemaExtensionsStyles();
  const [selectedSchemaExtension, setSelectedSchemaExtension] = React.useState<
    ISchemaExtension | undefined
  >(undefined);
  const [activeDrawer, setActiveDrawer] = React.useState<EActiveDrawer | null>(
    null,
  );
  const [refreshTrigger, setRefreshTrigger] = React.useState(0);
  const [isDeleteDialogOpen, setIsDeleteDialogOpen] = React.useState(false);
  const [isChangeStatusDialogOpen, setIsChangeStatusDialogOpen] =
    React.useState(false);
  const { logInfo } = useLogging();
  const { showSuccessToast} = useAppToast();
  const { title } = appGlobalState;

  // Derived state for drawer visibility
  const { isDrawerOpen, drawerMode, isViewerOpen, isResourceManagerOpen } =
    React.useMemo(
      () => ({
        isDrawerOpen:
          activeDrawer === EActiveDrawer.Create ||
          activeDrawer === EActiveDrawer.Edit,
        drawerMode:
          activeDrawer === EActiveDrawer.Edit
            ? ("edit" as const)
            : ("create" as const),
        isViewerOpen: activeDrawer === EActiveDrawer.View,
        isResourceManagerOpen:
          activeDrawer === EActiveDrawer.ResourceManager,
      }),
      [activeDrawer],
    );

  const canBeEdited = React.useMemo(
    () =>
      selectedSchemaExtension?.status === "InDevelopment" ||
      selectedSchemaExtension?.status === "Available" ||
      selectedSchemaExtension?.status === undefined,

    [selectedSchemaExtension],
  );

  const handleAddSchemaExtension = React.useCallback(() => {
    setSelectedSchemaExtension(undefined);
    setActiveDrawer(EActiveDrawer.Create);
  }, []);

  const handleSchemaExtensionSelect = React.useCallback(
    (schemaExtension: ISchemaExtension | undefined) => {
      setSelectedSchemaExtension(schemaExtension);

      // If a view/edit drawer is open, keep it open with the new selection
      // If no item is selected, close the view/edit drawer
      if (activeDrawer === EActiveDrawer.View || activeDrawer === EActiveDrawer.Edit) {
        if (!schemaExtension) {
          setActiveDrawer(null);
        }
      }
    },
    [activeDrawer],
  );

  const handleEditSchemaExtension = React.useCallback(
    (schemaExtension: ISchemaExtension) => {
      setSelectedSchemaExtension(schemaExtension);
      setActiveDrawer(EActiveDrawer.Edit);
    },
    [],
  );

  const handleDeleteSchemaExtension = React.useCallback(
    (schemaExtension: ISchemaExtension) => {
      logInfo("Delete schema extension:", JSON.stringify(schemaExtension));
     
        setActiveDrawer(null);
     
      setIsDeleteDialogOpen(true);
    },
    [],
  );

  const handleViewSchemaExtension = React.useCallback(
    (schemaExtension: ISchemaExtension) => {
      logInfo("View schema extension:", JSON.stringify(schemaExtension));
      setSelectedSchemaExtension(schemaExtension);
      setActiveDrawer(EActiveDrawer.View);
    },
    [logInfo],
  );

  const handleChangeSchemaStatus = React.useCallback(
    (schemaExtension: ISchemaExtension) => {
      logInfo("Change schema status:", JSON.stringify(schemaExtension));
      setSelectedSchemaExtension(schemaExtension);
      
        setActiveDrawer(null);
  
      setIsChangeStatusDialogOpen(true);
    },
    [logInfo],
  );

  const handleRefresh = React.useCallback(() => {
    setRefreshTrigger((prev) => prev + 1);
    // Clear selection when refreshing
    setSelectedSchemaExtension(undefined);
  }, []);

  const handleDrawerClose = React.useCallback(() => {
    setActiveDrawer(null);
  }, []);

  const handleViewerClose = React.useCallback(() => {
    setActiveDrawer(null);
  }, []);

  const handleOpenResourceManager = React.useCallback(() => {
    setActiveDrawer(EActiveDrawer.ResourceManager);
  }, []);

  const handleResourceManagerClose = React.useCallback(() => {
    setActiveDrawer(null);
  }, []);

  const handleDrawerSuccess = React.useCallback(() => {
    setActiveDrawer(null);
    setSelectedSchemaExtension(undefined);
    showSuccessToast("Operation completed successfully");
    // Trigger a refresh of the list view
    setRefreshTrigger((prev) => prev + 1);
  }, []);

  // Check if edit is disabled (Available status extensions can't be edited)
  const isEditDisabled = selectedSchemaExtension?.status === "Available";

  // Build toolbar items using the new ListToolbar
  const toolbarItems: IToolbarItem[] = React.useMemo(() => {
    return [
      // Add button
      {
        key: "add",
        label: "New",
        tooltip: "Create a new schema extension",
        icon: <Add20Regular />,
        onClick: handleAddSchemaExtension,
        appearance: "primary",
        group: "actions",
        dividerAfter: false,
      },
      // Edit button
      {
        key: "edit",
        label: "Edit",
        tooltip: !selectedSchemaExtension
          ? "Select a schema extension to edit"
          : isEditDisabled
            ? "Available schemas cannot be edited"
            : "Edit the selected schema extension",
        icon: <Edit20Regular />,
        onClick: () =>
          selectedSchemaExtension &&
          handleEditSchemaExtension(selectedSchemaExtension),
        disabled: !selectedSchemaExtension || isEditDisabled,
        group: "selection",
        appearance: "subtle",
      },
      // View button
      {
        key: "view",
        label: "View",
        tooltip: !selectedSchemaExtension
          ? "Select a schema extension to view"
          : "View the selected schema extension",
        icon: <Eye20Regular />,
        onClick: () =>
          selectedSchemaExtension &&
          handleViewSchemaExtension(selectedSchemaExtension),
        disabled: !selectedSchemaExtension,
        group: "selection",
        appearance: "subtle",
      },
      // Delete button
      {
        key: "delete",
        label: "Delete",
        tooltip: !selectedSchemaExtension
          ? "Select a schema extension to delete"
          : isEditDisabled
            ? "Available schemas cannot be deleted"
            : "Delete the selected schema extension",
        icon: <Delete20Regular />,
        onClick: () =>
          selectedSchemaExtension &&
          handleDeleteSchemaExtension(selectedSchemaExtension),
        disabled: !selectedSchemaExtension || isEditDisabled,
        group: "selection",
        appearance: "subtle",
      },
      // Change Status button
      {
        key: "changeStatus",
        label: "Change Status",
        tooltip: !selectedSchemaExtension
          ? "Select a schema extension to change status"
          : "Change the status of the selected schema extension",
        icon: <ArrowSync20Regular />,
        onClick: () =>
          selectedSchemaExtension &&
          handleChangeSchemaStatus(selectedSchemaExtension),
        disabled: !selectedSchemaExtension,
        group: "selection",
        appearance: "subtle",
      },
    ];
  }, [
    selectedSchemaExtension,
    isEditDisabled,
    handleAddSchemaExtension,
    handleEditSchemaExtension,
    handleViewSchemaExtension,
    handleDeleteSchemaExtension,
    handleChangeSchemaStatus,
  ]);

  // Far items (right side)
  const farItems: IToolbarItem[] = React.useMemo(() => {
    return [
      // Resource Manager button
      {
        key: "resourceManager",
        label: "Manage Resources",
        tooltip: "Manage schema extension attributes on resources",
        icon: <DatabaseSearch20Regular />,
        onClick: handleOpenResourceManager,
        group: "resources",
        appearance: "subtle" as const,
      },
      // Refresh button
      {
        key: "refresh",
        label: "Refresh",
        tooltip: "Refresh the list",
        icon: <ArrowClockwise20Regular />,
        onClick: handleRefresh,
        group: "refresh",
        appearance: "subtle" as const,
      },
    ];
  }, [handleOpenResourceManager, handleRefresh]);

  return (
    <div className={styles.appContainer}>
      <div className={styles.mainContent}>
        <Stack>
          <Stack gap="5px">
            <Subtitle1>{title}</Subtitle1>
            <InformationPanel />
          </Stack>

          <ListToolbar
            items={toolbarItems}
            farItems={farItems}
            ariaLabel="Schema Extensions Toolbar"
            showGroupDividers={true}
          />

          <Divider />
          <Stack className={styles.listViewContainer} >
            <SchemaExtensionsListView
              onSchemaExtensionSelect={handleSchemaExtensionSelect}
              onEdit={handleEditSchemaExtension}
              onDelete={handleDeleteSchemaExtension}
              onView={handleViewSchemaExtension}
              refreshTrigger={refreshTrigger}
              onCreateNew={handleAddSchemaExtension}
              onChangeStatus={handleChangeSchemaStatus}
            />
          </Stack>
        </Stack>
      </div>

      {/* Inline Drawers */}
      {isDrawerOpen && canBeEdited && (
        <SchemaExtensionDrawer
          isOpen={isDrawerOpen}
          onSchemaExtensionCreated={handleDrawerSuccess}
          onSchemaExtensionUpdated={handleDrawerSuccess}
          selectedSchemaExtension={selectedSchemaExtension}
          mode={drawerMode}
          onClose={handleDrawerClose}
        />
      )}

      {isViewerOpen && (
        <SchemaExtensionViewer
          isOpen={isViewerOpen}
          onClose={handleViewerClose}
          schemaExtension={selectedSchemaExtension}
        />
      )}

      <ResourceSchemaManager
        isOpen={isResourceManagerOpen}
        onClose={handleResourceManagerClose}
        onSuccess={(resourceId, data) => {
          logInfo(

            "Resource attributes updated:",
            JSON.stringify({ resourceId, data }),
          );
        }}
      />

      {/* Dialogs (remain as overlays) */}
      {isDrawerOpen && !canBeEdited && (
        <SchemaStatusRestrictionDialog
          isOpen={isDrawerOpen}
          onDismiss={handleDrawerClose}
          schemaExtension={selectedSchemaExtension!}
        />
      )}
      {isDeleteDialogOpen && selectedSchemaExtension && (
        <DeleteSchemaExtension
          isOpen={isDeleteDialogOpen}
          onDismiss={() => setIsDeleteDialogOpen(false)}
          schemaExtension={selectedSchemaExtension!}
          onDeleteSuccess={() => {
            setIsDeleteDialogOpen(false);
            setSelectedSchemaExtension(undefined);
            setRefreshTrigger((prev) => prev + 1);
          }}
        />
      )}
      {isChangeStatusDialogOpen && selectedSchemaExtension && (
        <ChangeSchemaStatus
          isOpen={isChangeStatusDialogOpen}
          onDismiss={() => setIsChangeStatusDialogOpen(false)}
          schemaExtension={selectedSchemaExtension!}
          onStatusChangeSuccess={() => {
            setIsChangeStatusDialogOpen(false);
            setSelectedSchemaExtension(undefined);
            setRefreshTrigger((prev) => prev + 1);
          }}
        />
      )}
    </div>
  );
};

export default ManageSchemaExtensionsControl;
