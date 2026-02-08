import { css } from '@emotion/css';
import { tokens } from '@fluentui/react-components';

export interface IUserPickerStyles {
  tagPickerControl: string;
  tagPickerInput: string;
  container: string;
  pickerList: string;
  optionContent: string;
  optionSecondary: string;
}

export const useUserPickerStyles = (): IUserPickerStyles => {
  return {
    container: css({
      display: "flex",
      flexDirection: "column",
      gap: tokens.spacingVerticalS,
      position: "relative",
    }),

    pickerList: css({
      maxHeight: 300,
      overflowY: "auto",
      backgroundColor: tokens.colorNeutralBackground1,
      border: `1px solid ${tokens.colorNeutralStroke2}`,
      borderRadius: tokens.borderRadiusMedium,
      boxShadow: tokens.shadow16,
      scrollbarColor: `${tokens.colorBrandForeground1} ${tokens.colorNeutralBackground1}`,
      scrollbarWidth: "thin",
      "&::-webkit-scrollbar-track": {
        borderRadius: 10,
        borderWidth: 1,
      },
      "&::-webkit-scrollbar-thumb": {
        borderRadius: 10,
        borderWidth: 1,
      },
      "&::-webkit-scrollbar": {
        height: 7,
        width: 7,
      },
    }),

    optionContent: css({
      display: "flex",
      flexDirection: "column",
      gap: tokens.spacingVerticalXXS,
    }),

    optionSecondary: css({
      fontSize: tokens.fontSizeBase200,
      color: tokens.colorNeutralForeground3,
      fontWeight: tokens.fontWeightRegular as unknown as number,
    }),

    tagPickerControl: css({
      border: `1px solid ${tokens.colorNeutralStroke2}`,
      borderRadius: 50,
      backgroundColor: tokens.colorNeutralBackground1,
      padding: "0px 10px",
      boxShadow: "none",
      transition: "all 0.2s ease",
      "&::after": {
        display: "none !important",
      },
      "&:hover": {
        borderColor: tokens.colorBrandStroke1,
      },
      "&:focus-within": {
        borderColor: tokens.colorBrandStroke1,
        boxShadow: `0 0 0 2px ${tokens.colorBrandBackground2}`,
      },
    }),

    tagPickerInput: css({
      backgroundColor: "transparent",
      border: "none",
      outline: "none",
      fontSize: tokens.fontSizeBase300,
      "&::placeholder": {
        color: tokens.colorNeutralForeground4,
      },
      "&:focus": {
        outline: "none",
      },
    }),
  };
};
