import * as React from 'react';
import "@pnp/polyfill-ie11";
import './ShowFoldersPermissionsWise.module.scss';
import { IShowFoldersPermissionsWiseProps } from './IShowFoldersPermissionsWiseProps';
import { IFoldersState, IFolderItem } from "./IShowFoldersState";
import { sp, PermissionKind } from "@pnp/sp/presets/all";
import { Shimmer } from 'office-ui-fabric-react/lib/Shimmer';
import { IDocumentLibraryInformation } from "@pnp/sp/sites";
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Icon, MessageBar } from 'office-ui-fabric-react';

export default class ShowFoldersPermissionsWise extends React.Component<IShowFoldersPermissionsWiseProps, IFoldersState> {

  constructor(props) {
    super(props);

    this.state = {
      foldersData: [],
      isLoaded: false
    };
  }

  //This method gets called in the page load.
  public componentDidMount() {
    try {
      this.getAllLibrariesFoldersData();
    }
    catch (error) {
      console.log(error);
      this.setState({
        isLoaded: true
      });
    }
  }

  //This method fetches all the accessible folders across all libraries present in a site.
  private async getAllLibrariesFoldersData() {

    try {

      let folderItems: IFolderItem[] = new Array();

      const docLibs: any = await sp.site.getDocumentLibraries(this.props.siteURL);

      //we got the array of document libraries
      docLibs.results.forEach((docLib: IDocumentLibraryInformation) => {

        //Do not show the Shared Documents library in the web part.
        if (docLib.ServerRelativeUrl.split("/").pop() !== "Shared Documents") {
          //Check if current user has library level permissions
          let hasLibraryPermissionsPromise = this.checkUserLibraryPermissions(docLib);

          hasLibraryPermissionsPromise.then((hasLibraryPermissions: boolean) => {

            //If current user has library level permissions then show the document library name directly
            if (hasLibraryPermissions == true) {

              let folderItem: IFolderItem = {
                FolderName: docLib.Title,
                FolderLink: docLib.ServerRelativeUrl,
              };
              folderItems.push(folderItem);

              //Call to sort the library and folder items and set state
              this.sortAndSetState(folderItems);
            }
            else {

              //otherwise parse each library to fetch folders to show each accessible folders.
              let foldersDataPromise = this.getFoldersData(docLib);

              foldersDataPromise.then((folders: any[]) => {

                folders.forEach((folder) => {
                  let folderItem: IFolderItem = {
                    FolderName: folder.FileLeafRef,
                    FolderLink: folder.FileRef
                  };

                  let result: any = folderItems.filter(fItem =>
                    folder.FileRef.indexOf(fItem.FolderLink) > -1 &&
                    folder.FileLeafRef.indexOf(fItem.FolderName) == -1

                  );

                  if (result.length == 0)
                    folderItems.push(folderItem);

                });

                //Call to sort the library and folder items and set state
                this.sortAndSetState(folderItems);
              });
            }

          });
        }
      });
    }
    catch (error) {
      console.log(error);
      this.setState({
        isLoaded: true
      });
    }

  }

  //This method sorts the library & folder items and updates the react state.
  private sortAndSetState(folderItems: IFolderItem[]) {

    folderItems.sort((a, b) => {
      if (a.FolderName < b.FolderName) { return -1; }
      if (a.FolderName > b.FolderName) { return 1; }
      return 0;
    });

    this.setState({
      foldersData: folderItems,
      isLoaded: true
    });

  }

  //This method gets all the folders available in the passed document library.
  private async getFoldersData(docLib: any) {

    let folders: any[] = await sp.web.lists.getByTitle(docLib.Title)
      .items
      .filter('FSObjType eq 1')
      .select('FileLeafRef', 'FileRef')
      .get();

    return folders;

  }

  //Check if current user has permissions and return the promise for the same.
  private async checkUserLibraryPermissions(docLib: any) {

    const hasLibraryPermissionsPromise = await sp.web.lists.getByTitle(docLib.Title).currentUserHasPermissions(PermissionKind.ViewListItems);
    return hasLibraryPermissionsPromise;

  }

  //Creates the html JSX based on the folders data loaded from SharePoint libraries within a site.
  public createJSXForFoldersAndLibrary(): any {

    let linkIconName = this.props.iconName !== null && this.props.iconName !== "" && this.props.iconName !== undefined ? this.props.iconName : "Globe";

    if (this.state.foldersData.length > 0) {
      return this.state.foldersData.map((folderItem) => {
        return (
          <React.Fragment>
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 ms-xl12 ms-xxl12 ms-xxxl12 libTitle">
              <div className="folderIcon">
                <Icon iconName={linkIconName} ariaLabel={linkIconName}></Icon>
              </div>
              <div className="folderLink">
                <a href={window.location.origin + folderItem.FolderLink} className="anchorTag">
                  {folderItem.FolderName}
                </a>
              </div>
            </div>
          </React.Fragment>

        );
      });

    }

  }

  public componentDidUpdate() {
    try {
      this._colorConfig();
    }
    catch (error) {
      console.log(error);
      this.setState({
        isLoaded: true
      });
    }
  }

  //Apply the selected colors to the UI elements including web part title, background and icon.
  private _colorConfig() {
    let webParttitleColor: any = document.querySelectorAll(".webPartTitleSection div[class^='webPartTitle_']")[0];
    if (webParttitleColor !== undefined && webParttitleColor !== null)
      webParttitleColor.style.color = this.props.titleColor;

    let webPartbackgroundColor: any = document.querySelectorAll(".showFoldersPermissionsWiseWrapper")[0];
    if (webPartbackgroundColor !== undefined && webPartbackgroundColor !== null)
      webPartbackgroundColor.style.background = this.props.backgroundColor;

    let titleColor: any = document.querySelectorAll(".folderLink a");
    let iconBackgroundColor: any = document.querySelectorAll(".folderIcon i");
    const elemLength: number = iconBackgroundColor.length;

    if (titleColor !== undefined && iconBackgroundColor !== null && elemLength > 0) {
      for (let i = 0; i < elemLength; i++) {
        titleColor[i].style.color = this.props.titleColor;
        iconBackgroundColor[i].style.color = this.props.iconBackgroundColor;
      }
    }

  }

  //Show shimmers till the data gets loaded on the screen
  public getLoadingShimmers(): any {

    return (<div>
      <Shimmer /><br />
      <Shimmer /><br />
      <Shimmer /><br />
      <Shimmer /><br />
      <Shimmer /><br />
      <Shimmer /><br />
      <Shimmer /><br />
      <Shimmer />
    </div>
    );

  }

  //Rendering the html
  public render(): React.ReactElement<IShowFoldersPermissionsWiseProps> {

    const webpartTitle = <WebPartTitle displayMode={this.props.displayMode}
      title={this.props.title}
      updateProperty={this.props.updateProperty} />;

    return (
      <div className="showFoldersPermissionsWiseWrapper" >

        <div className="webPartTitleSection">
          {webpartTitle}
        </div>

        <div className="showFoldersData">
          {
            this.state.isLoaded == true ?
              this.state.foldersData.length <= 0
                ?
                <div className="emptySection">
                  <div className="emptyTextSection">
                    <MessageBar>{this.props.noFoldersFoundMessage !== null && this.props.noFoldersFoundMessage !== "" && this.props.noFoldersFoundMessage !== undefined ? this.props.noFoldersFoundMessage : "No data found."}</MessageBar>
                  </div>
                </div>
                :
                <div className="ms-Grid custom-row subTitle" dir="ltr">
                  {this.createJSXForFoldersAndLibrary()}
                </div>
              :
              this.getLoadingShimmers()
          }
        </div>
      </div>

    );
  }
}
