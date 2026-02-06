import * as React from "react";
import {
  Toolbar,
  ToolbarButton,
  ToolbarDivider,
  ToolbarGroup,
  Tooltip,
  Badge,
  Text,
} from "@fluentui/react-components";
import { useListToolbarStyles } from "./useListToolbarStyles";
import { IToolbarItem } from "./IToolbarItem";
import { IListToolbarProps } from "./IListToolbarProps";

/**
 * Renders a single toolbar item
 */
const ToolbarItemRenderer: React.FC<{
  item: IToolbarItem;
  isLoading?: boolean;
}> = ({ item, isLoading }) => {
  // Skip if not visible
  if (item.visible === false) {
    return null;
  }

  // Use custom render if provided
  if (item.onRender) {
    return item.onRender();
  }

  const button = (
    <ToolbarButton
      key={item.key}
      aria-label={item.ariaLabel || item.tooltip || item.label}
      icon={item.icon}
      onClick={item.onClick}
      disabled={item.disabled || isLoading}
      appearance={item.appearance}
    >
      {item.label}
    </ToolbarButton>
  );

  // Wrap with tooltip if provided
  if (item.tooltip) {
    return (
      <Tooltip content={item.tooltip} relationship="label">
        {button}
      </Tooltip>
    );
  }

  return button;
};

/**
 * Groups toolbar items by their group property
 */
const groupItems = (items: IToolbarItem[]): Map<string, IToolbarItem[]> => {
  const groups = new Map<string, IToolbarItem[]>();
  
  items.forEach((item) => {
    const groupName = item.group || "default";
    if (!groups.has(groupName)) {
      groups.set(groupName, []);
    }
    groups.get(groupName)!.push(item);
  });
  
  return groups;
};

/**
 * A flexible toolbar component that accepts a list of toolbar items
 * with support for grouping, far items, and custom rendering.
 */
export const ListToolbar: React.FunctionComponent<IListToolbarProps> = ({
  items,
  farItems = [],
  isLoading = false,
  ariaLabel = "Toolbar",
  totalCount,
  className,
  showGroupDividers = true,
}) => {
  const styles = useListToolbarStyles();

  // Separate items into regular and far items
  const regularItems = React.useMemo(() => {
    return items.filter((item) => !item.isFarItem && item.visible !== false);
  }, [items]);

  const allFarItems = React.useMemo(() => {
    const itemsFarItems = items.filter(
      (item) => item.isFarItem && item.visible !== false
    );
    const visibleFarItems = farItems.filter((item) => item.visible !== false);
    return [...itemsFarItems, ...visibleFarItems];
  }, [items, farItems]);

  // Group regular items
  const groupedItems = React.useMemo(() => {
    return groupItems(regularItems);
  }, [regularItems]);

  // Group far items
  const groupedFarItems = React.useMemo(() => {
    return groupItems(allFarItems);
  }, [allFarItems]);

  /**
   * Renders a group of items with optional dividers
   */
  const renderGroup = (
    groupItems: IToolbarItem[],
    groupIndex: number,
    isLastGroup: boolean
  ): React.ReactNode[] => {
    const elements: React.ReactNode[] = [];

    groupItems.forEach((item, itemIndex) => {
      // Add divider before if specified
      if (item.dividerBefore) {
        elements.push(<ToolbarDivider key={`${item.key}-divider-before`} />);
      }

      elements.push(
        <ToolbarItemRenderer
          key={item.key}
          item={item}
          isLoading={isLoading}
        />
      );

      // Add divider after if specified
      if (item.dividerAfter) {
        elements.push(<ToolbarDivider key={`${item.key}-divider-after`} />);
      }
    });

    // Add group divider if not the last group and showGroupDividers is true
    if (showGroupDividers && !isLastGroup) {
      elements.push(<ToolbarDivider key={`group-divider-${groupIndex}`} />);
    }

    return elements;
  };

  return (
    <Toolbar aria-label={ariaLabel} className={`${styles.toolbar} ${className || ""}`}>
      {/* Left side - regular items */}
      <ToolbarGroup className={styles.leftGroup}>
        {Array.from(groupedItems.entries()).map(
          ([_groupName, groupItemsList], groupIndex) =>
            renderGroup(
              groupItemsList,
              groupIndex,
              groupIndex === groupedItems.size - 1
            )
        )}
        
        {/* Total count badge */}
        {totalCount !== undefined && totalCount > 0 && (
          <Badge
            appearance="filled"
            color="informative"
            className={styles.countBadge}
          >
            <Text size={200}>{totalCount} items</Text>
          </Badge>
        )}
      </ToolbarGroup>

      {/* Right side - far items */}
      {allFarItems.length > 0 && (
        <ToolbarGroup className={styles.rightGroup}>
          {Array.from(groupedFarItems.entries()).map(
            ([_groupName, groupItemsList], groupIndex) =>
              renderGroup(
                groupItemsList,
                groupIndex,
                groupIndex === groupedFarItems.size - 1
              )
          )}
        </ToolbarGroup>
      )}
    </Toolbar>
  );
};

export default ListToolbar;
