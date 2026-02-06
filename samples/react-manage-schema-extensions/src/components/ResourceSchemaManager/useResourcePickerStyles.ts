import { css } from "@emotion/css";
import { tokens } from "@fluentui/react-components";

export interface IResourcePickerStyles {
  container: string;
  tagPickerStyles: string;
  pickerList: string;
  optionContent: string;
  optionSecondary: string;
}

export const useResourcePickerStyles = (): IResourcePickerStyles => {
  return {
    container: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalS};
      position: relative;
    `,

    pickerList: css`
      max-height: 300px;
      overflow-y: auto;
      background-color: ${tokens.colorNeutralBackground1};
      border: 1px solid ${tokens.colorNeutralStroke2};
      border-radius: ${tokens.borderRadiusMedium};
      box-shadow: ${tokens.shadow16};
      scrollbar-color: ${tokens.colorBrandForeground1} ${tokens.colorNeutralBackground1};
      scrollbar-width: thin;

      &::-webkit-scrollbar-track {
        border-radius: 10px;
        border-width: 1px;
      }

      &::-webkit-scrollbar-thumb {
        border-radius: 10px;
        border-width: 1px;
      }

      &::-webkit-scrollbar {
        height: 7px;
        width: 7px;
      }
    `,

    optionContent: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalXXS};
    `,

    optionSecondary: css`
      font-size: ${tokens.fontSizeBase200};
      color: ${tokens.colorNeutralForeground3};
      font-weight: ${tokens.fontWeightRegular};
    `,

    tagPickerStyles: css`
      /* Target the root element and its direct children */
      > div {
        border: 1px solid ${tokens.colorNeutralStroke2};
        border-radius: 50px;
        background-color: ${tokens.colorNeutralBackground1};
        padding: 0px 10px;
        box-shadow: none;
        transition: all 0.2s ease;
        position: relative;

        /* Remove default focus styling */
        &::after {
          display: none !important;
        }

        &:hover {
          border-color: ${tokens.colorBrandStroke1};
        }

        &:focus-within {
          border-color: ${tokens.colorBrandStroke1};
          box-shadow: 0 0 0 2px ${tokens.colorBrandBackground2};
        }
      }

      /* Style the input */
      input {
        background-color: transparent;
        border: none;
        outline: none;
        font-size: ${tokens.fontSizeBase300};

        &::placeholder {
          color: ${tokens.colorNeutralForeground4};
        }
      }

      /* Style the tags */
      [data-fui-focus-visible] {
        outline: none;
      }
    `,
  };
};
