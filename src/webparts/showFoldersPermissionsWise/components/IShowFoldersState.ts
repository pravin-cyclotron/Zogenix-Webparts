export interface IFoldersState {
  foldersData: IFolderItem[];
  isLoaded: boolean;
}

export interface IFolderItem {
  FolderName: string;
  FolderLink: string;
}
