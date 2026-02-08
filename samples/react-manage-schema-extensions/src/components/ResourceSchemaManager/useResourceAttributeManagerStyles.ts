import { css } from "@emotion/css";
import { tokens } from "@fluentui/react-components";

export interface IResourceAttributeManagerStyles {
  propertiesContainer: string;
  scrollableContent: string;
  root: string;
  innerContent: string;
}

export const useResourceAttributeManagerStyles =
  (hasMultipleTargetTypes: boolean): IResourceAttributeManagerStyles => {
    return {
      root: css({
        display: "flex",
        flexDirection: "column",
        flex: 1,
        minHeight: 0,
        overflow: "hidden",
      }),
      innerContent: css({
        display: "flex",
        flexDirection: "column",
        flex: 1,
        minHeight: 0,
        overflow: "hidden",
      }),
      propertiesContainer: css({
        flex: 1,
        minHeight: 0,
        overflowY: "auto",
        scrollbarColor: `${tokens.colorBrandForeground1} transparent`,
        scrollbarWidth: "thin",
        "&::-webkit-scrollbar": {
          width: 8,
        },
        "&::-webkit-scrollbar-track": {
          borderRadius: 10,
          backgroundColor: "transparent",
        },
        "&::-webkit-scrollbar-thumb": {
          borderRadius: 10,
          backgroundColor: tokens.colorBrandForeground1,
        },
      }),
         scrollableContent: css({
        flex: 1,
        overflowY: "auto",
        paddingRight: tokens.spacingHorizontalXS,
      }),
    };
  };
