import { DisplayMode } from "@microsoft/sp-core-library";

export interface IShowFoldersPermissionsWiseProps {
  siteURL: string;
  title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  iconName: string;
  noFoldersFoundMessage: string;
  titleColor: string;
  backgroundColor: string;
  iconBackgroundColor: string;
}
