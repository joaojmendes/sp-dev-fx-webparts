import { css } from '@emotion/css';
import { tokens } from '@fluentui/react-components';

export interface IUserPickerStyles {
  tagPickerStyles: string;
  container: string;
  pickerList: string;
  optionContent: string;
  optionSecondary: string;
}

export const useUserPickerStyles = (): IUserPickerStyles => {
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
  
      font-weight: ${tokens.fontWeightRegular};
    `,
    
    tagPickerStyles: css`
      
      /* Target the root element and its direct children */
      > div {
        border: 1px solid ${tokens.colorNeutralStroke2};
        border-radius: 50px; /* Fully circular/pill shape */
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
          outline: none;
          border-color: ${tokens.colorBrandStroke1};
        }

        &:disabled {
          opacity: 0.6;
          cursor: not-allowed;
          border-color: ${tokens.colorNeutralStroke3};
          background-color: ${tokens.colorNeutralBackground3};
        }
      }

      /* Remove any internal focus indicators from all elements */
      * {
        &:focus {
          outline: none !important;
          box-shadow: none !important;
        }
      }

      /* Style the tag group container using attribute selectors */
      [role="group"] {
        display: flex;
        flex-wrap: wrap;
        gap: ${tokens.spacingHorizontalXS};
        margin-right: ${tokens.spacingHorizontalS};
      }

      /* Style individual tags using proper selectors */
      [role="button"][aria-label*="Remove"] {
        margin: 0;
      }
    `
  };
};
