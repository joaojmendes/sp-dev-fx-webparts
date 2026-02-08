import { css } from "@emotion/css";
import { tokens } from "@fluentui/react-components";

export interface ISchemaExtensionViewerStyles {
  drawerRoot: string;
  drawerFooter: string;
  drawerContent: string;
  scrollableContent: string;
  noDataContainer: string;
  sectionIcon: string;
  sectionContent: string;
  readOnlyField: string;
  fieldIcon: string;
  emptyPropertiesMessage: string;
  iconStyles: string;
  targetTypesContainer: string;
  inputDisabledColor: string;
}

export const useSchemaExtensionViewerStyles =
  (): ISchemaExtensionViewerStyles => {
    return {
      drawerRoot: css({
        backgroundColor: tokens.colorNeutralBackground1,
      
        borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
      }),
      drawerFooter: css({
        backgroundColor: tokens.colorNeutralBackground1,
      }),
      inputDisabledColor: css({
        backgroundColor: tokens.colorNeutralBackground3,
        color: tokens.colorNeutralForeground1Selected,
      }),
      drawerContent: css({
        display: "flex",
        flexDirection: "column",
        gap: tokens.spacingVerticalM,
        height: "100%",
        paddingTop: tokens.spacingVerticalL,
        paddingLeft: tokens.spacingHorizontalM,
        paddingRight: tokens.spacingHorizontalM,
      }),

      scrollableContent: css({
        flex: 1,
        overflowY: "auto",
        paddingRight: tokens.spacingHorizontalXS,
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

      targetTypesContainer: css({
        display: "grid",
        gridTemplateColumns: "repeat(4, auto)",
        gap: tokens.spacingHorizontalS,
        width: "100%",
        backgroundColor: tokens.colorNeutralBackground3,
        padding: "5px",
      }),

      iconStyles: css({
        color: tokens.colorBrandBackground,
      }),

      emptyPropertiesMessage: css({
        padding: tokens.spacingVerticalL,
        textAlign: "center",
        fontStyle: "italic",
        color: tokens.colorNeutralForeground3,
      }),

      noDataContainer: css({
        padding: tokens.spacingVerticalXXL,
        alignItems: "center",
        justifyContent: "center",
        minHeight: "200px",
      }),

      sectionIcon: css({
        fontSize: "20px",
        color: tokens.colorBrandForeground1,
      }),

      sectionContent: css({
        marginTop: tokens.spacingVerticalM,
      }),

      readOnlyField: css({
        padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
        backgroundColor: tokens.colorNeutralBackground2,
        borderRadius: tokens.borderRadiusMedium,
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        minHeight: "32px",
        display: "flex",
        alignItems: "center",
      }),

      fieldIcon: css({
        fontSize: "16px",
        color: tokens.colorNeutralForeground3,
      }),
    };
  };
