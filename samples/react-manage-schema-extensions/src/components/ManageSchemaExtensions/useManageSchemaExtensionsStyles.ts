import { css } from '@emotion/css';
import { tokens } from '@fluentui/react-components';

export interface ManageSchemaExtensionsStyles {
  title: string;
  description: string;
  toolbar: string;
  content: string;
  sharePointHost: string;
  otherHosts: string;
  appContainer: string;
  mainContent: string;
    listViewContainer: string;
}

export const useManageSchemaExtensionsStyles = (): ManageSchemaExtensionsStyles => {
  return {
    listViewContainer: css({
      paddiingTop: tokens.spacingVerticalXXL,
        overflowX: "auto",
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
    title: css({
      fontSize: tokens.fontSizeHero800,
      fontWeight: tokens.fontWeightSemibold,
      color: tokens.colorNeutralForeground1,
      margin: 0,
      marginBottom: tokens.spacingVerticalS,
    }),
    description: css({
      fontSize: tokens.fontSizeBase300,
      color: tokens.colorNeutralForeground2,
      margin: 0,
    }),
    toolbar: css({
      marginBottom: tokens.spacingVerticalM,
    }),
    content: css({
      flex: 1,
      overflow: 'hidden',
    }),
    sharePointHost: css({
      backgroundColor: 'transparent',
    }),
    otherHosts: css({
      backgroundColor: tokens.colorNeutralBackground1,
      height: 'calc(100vh - 64px)',
      padding: tokens.spacingVerticalL,
      paddingLeft: tokens.spacingHorizontalL,
      paddingRight: tokens.spacingHorizontalL,
    }),
    appContainer: css({
      display: 'flex',
      flexDirection: 'row',
      height: '100%',
      overflow: 'hidden',
    }),
    mainContent: css({
      flex: 1,
      minWidth: 0,
      overflow: 'auto',
    }),
  };
};