export interface IFoldersState {
  foldersData: IFolderItem[];
  isLoaded: boolean;
  foldersHTML: any;
}

export interface IFolderItem {
  DocumentLibrary: string;
  FolderName: string;
  FolderLink: Text;
}
