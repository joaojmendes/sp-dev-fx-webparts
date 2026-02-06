import { css } from "@emotion/css";
import { tokens } from "@fluentui/react-components";

export interface IListToolbarStyles {
  toolbar: string;
  leftGroup: string;
  rightGroup: string;
  countBadge: string;
  divider: string;
}

export const useListToolbarStyles = (): IListToolbarStyles => {
  return {
    toolbar: css`
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: ${tokens.spacingVerticalS} ${tokens.spacingHorizontalM};
      background-color: ${tokens.colorNeutralBackground1};
      border-bottom: 1px solid ${tokens.colorNeutralStroke2};
      min-height: 48px;
      flex-wrap: wrap;
      gap: ${tokens.spacingVerticalXS};
    `,

    leftGroup: css`
      display: flex;
      align-items: center;
      gap: ${tokens.spacingHorizontalXS};
      flex-wrap: wrap;
    `,

    rightGroup: css`
      display: flex;
      align-items: center;
      gap: ${tokens.spacingHorizontalXS};
      flex-wrap: wrap;
    `,

    countBadge: css`
      margin-left: ${tokens.spacingHorizontalS};
    `,

    divider: css`
      height: 24px;
      margin: 0 ${tokens.spacingHorizontalXS};
    `,
  };
};
