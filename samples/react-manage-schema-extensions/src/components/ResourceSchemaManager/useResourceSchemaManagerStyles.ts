import { css } from "@emotion/css";
import { tokens } from "@fluentui/react-components";

export interface IResourceSchemaManagerStyles {
  container: string;
  header: string;
  title: string;
  subtitle: string;
  form: string;
  section: string;
  sectionTitle: string;
  dropdownContainer: string;
  resourceInfo: string;
  attributesContainer: string;
  attributeGroup: string;
  attributeRow: string;
  inputGrid: string;
  buttonGroup: string;
  loadingContainer: string;
  errorContainer: string;
  successContainer: string;
  emptyState: string;
  infoCard: string;
  schemaInfo: string;
  propertyBadge: string;
  targetTypeBadge: string;
  resourceCard: string;
  pickerList: string;
  optionContent: string;
  optionSecondary: string;
}

export const useResourceSchemaManagerStyles = (): IResourceSchemaManagerStyles => {
  return {
    container: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalL};
      padding: ${tokens.spacingVerticalL};
      background-color: ${tokens.colorNeutralBackground1};
      border-radius: ${tokens.borderRadiusMedium};
      border: 1px solid ${tokens.colorNeutralStroke2};
      box-shadow: ${tokens.shadow4};
      min-height: 400px;
    `,

    header: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalS};
      margin-bottom: ${tokens.spacingVerticalM};
      padding-bottom: ${tokens.spacingVerticalM};
      border-bottom: 1px solid ${tokens.colorNeutralStroke3};
    `,

    title: css`
      font-size: ${tokens.fontSizeBase500};
      font-weight: ${tokens.fontWeightSemibold};
      color: ${tokens.colorNeutralForeground1};
      margin: 0;
    `,

    subtitle: css`
      font-size: ${tokens.fontSizeBase300};
      color: ${tokens.colorNeutralForeground2};
      margin: 0;
    `,

    form: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalL};
    `,

    section: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalM};
      padding: ${tokens.spacingVerticalM};
      background-color: ${tokens.colorNeutralBackground2};
      border-radius: ${tokens.borderRadiusMedium};
      border: 1px solid ${tokens.colorNeutralStroke3};
    `,

    sectionTitle: css`
      font-size: ${tokens.fontSizeBase400};
      font-weight: ${tokens.fontWeightSemibold};
      color: ${tokens.colorNeutralForeground1};
      margin: 0;
      display: flex;
      align-items: center;
      gap: ${tokens.spacingHorizontalS};
    `,




    attributesContainer: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalM};
    `,

    attributeGroup: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalS};
    `,

    attributeRow: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalXS};
    `,

    inputGrid: css`
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: ${tokens.spacingHorizontalM};
    `,

    buttonGroup: css`
      display: flex;
      gap: ${tokens.spacingHorizontalM};
      justify-content: flex-end;
      margin-top: ${tokens.spacingVerticalL};
      padding-top: ${tokens.spacingVerticalM};
      border-top: 1px solid ${tokens.colorNeutralStroke3};
    `,

    loadingContainer: css`
      display: flex;
      align-items: center;
      justify-content: center;
      padding: ${tokens.spacingVerticalXXL};
      gap: ${tokens.spacingHorizontalM};
      flex-direction: column;
    `,

    errorContainer: css`
      padding: ${tokens.spacingVerticalM};
      background-color: ${tokens.colorPaletteRedBackground1};
      border: 1px solid ${tokens.colorPaletteRedBorder1};
      border-radius: ${tokens.borderRadiusMedium};
      color: ${tokens.colorPaletteRedForeground1};
    `,

    successContainer: css`
      padding: ${tokens.spacingVerticalM};
      background-color: ${tokens.colorPaletteGreenBackground1};
      border: 1px solid ${tokens.colorPaletteGreenBorder1};
      border-radius: ${tokens.borderRadiusMedium};
      color: ${tokens.colorPaletteGreenForeground1};
    `,

    emptyState: css`
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: ${tokens.spacingVerticalXXL};
      gap: ${tokens.spacingVerticalM};
      color: ${tokens.colorNeutralForeground3};
      text-align: center;
    `,

    infoCard: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalS};
      padding: ${tokens.spacingVerticalM};
      background-color: ${tokens.colorNeutralBackground3};
      border-radius: ${tokens.borderRadiusMedium};
      border: 1px solid ${tokens.colorNeutralStroke3};
    `,

    schemaInfo: css`
      display: flex;
      flex-wrap: wrap;
      gap: ${tokens.spacingHorizontalS};
      align-items: center;
    `,

    propertyBadge: css`
      display: inline-flex;
      align-items: center;
      gap: ${tokens.spacingHorizontalXS};
      padding: ${tokens.spacingVerticalXXS} ${tokens.spacingHorizontalS};
      background-color: ${tokens.colorBrandBackground2};
      color: ${tokens.colorBrandForeground2};
      border-radius: ${tokens.borderRadiusCircular};
      font-size: ${tokens.fontSizeBase200};
      font-weight: ${tokens.fontWeightMedium};
    `,

    targetTypeBadge: css`
      display: inline-flex;
      align-items: center;
      gap: ${tokens.spacingHorizontalXS};
      padding: ${tokens.spacingVerticalXXS} ${tokens.spacingHorizontalS};
      background-color: ${tokens.colorPalettePurpleBackground2};
      color: ${tokens.colorPalettePurpleForeground2};
      border-radius: ${tokens.borderRadiusCircular};
      font-size: ${tokens.fontSizeBase200};
      font-weight: ${tokens.fontWeightMedium};
    `,

    resourceCard: css`
      display: flex;
      align-items: center;
      gap: ${tokens.spacingHorizontalM};
      padding: ${tokens.spacingVerticalM};
      background-color: ${tokens.colorNeutralBackground1};
      border: 1px solid ${tokens.colorNeutralStroke2};
      border-radius: ${tokens.borderRadiusMedium};

      &:hover {
        background-color: ${tokens.colorNeutralBackground1Hover};
        border-color: ${tokens.colorBrandStroke1};
      }
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

    dropdownContainer: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalS};
    `,

    resourceInfo: css`
      display: flex;
      flex-direction: column;
      gap: ${tokens.spacingVerticalXS};
    `,
  };
};
