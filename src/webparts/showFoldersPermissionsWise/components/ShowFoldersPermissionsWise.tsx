import * as React from 'react';
import './ShowFoldersPermissionsWise.module.scss';
import { IShowFoldersPermissionsWiseProps } from './IShowFoldersPermissionsWiseProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IFoldersState, IFolderItem } from "./IShowFoldersState";
import { sp } from "@pnp/sp/presets/all";
import { Shimmer } from 'office-ui-fabric-react/lib/Shimmer';
import { IDocumentLibraryInformation } from "@pnp/sp/sites";
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';

export default class ShowFoldersPermissionsWise extends React.Component<IShowFoldersPermissionsWiseProps, IFoldersState> {

  constructor(props) {
    super(props);

    this.state = {
      foldersData: [],
      isLoaded: false,
      foldersHTML: null
    };
  }

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

  private async getAllLibrariesFoldersData() {

    try {

      let folderItems: IFolderItem[] = new Array();

      const docLibs: any = await sp.site.getDocumentLibraries(this.props.siteURL);

      //we got the array of document library information
      docLibs.results.forEach((docLib: IDocumentLibraryInformation) => {
        // parse each library to fetch folders

        let foldersDataPromise = this.getFoldersData(docLib);

        foldersDataPromise.then((folders: any[]) => {

          folders.forEach((folder) => {
            let folderItem: IFolderItem = {
              DocumentLibrary: docLib.Title,
              FolderName: folder.FileLeafRef,
              FolderLink: folder.FileRef
            };

            let result: any = folderItems.filter(fItem => folder.FileRef.indexOf(fItem.FolderLink) > -1);

            if (result.length == 0)
              folderItems.push(folderItem);

          });

          this.setState({
            foldersData: folderItems,
            isLoaded: true
          });
        });
      });
    }
    catch (error) {
      console.log(error);
      this.setState({
        isLoaded: true
      });
    }

  }

  private async getFoldersData(docLib: any) {

    let folders: any[] = await sp.web.lists.getByTitle(docLib.Title)
      .items
      .filter('FSObjType eq 1')
      .select('FileLeafRef', 'FileRef')
      .get();

    return folders;

  }

  public createJSXForFoldersAndLibrary(): any {

    if (this.state.foldersData.length > 0) {
      return this.state.foldersData.map((folderItem) => {
        return (

          <React.Fragment>

            <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg4 ms-xl4 ms-xxl4 ms-xxxl4 libTitle">
              {folderItem.DocumentLibrary}
            </div>
            <div className="ms-Grid-col ms-sm12 ms-md8 ms-lg8 ms-xl8 ms-xxl8 ms-xxxl8 libTitle">
              <a href={window.location.origin + folderItem.FolderLink} className="anchorTag" target="_blank" data-interception="off">
                {folderItem.FolderName}
              </a>
            </div>
          </React.Fragment>
        );
      });

    }

  }

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

          <div className="ms-Grid custom-row subTitle" dir="ltr">
            <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg4 ms-xl4 ms-xxl4 ms-xxxl4 libTitleHeader">
              Document Library
                  </div>
            <div className="ms-Grid-col ms-sm12 ms-md8 ms-lg8 ms-xl8 ms-xxl8 ms-xxxl8 libTitleHeader">
              Folder Link
                  </div>
          </div>

          {
            this.state.isLoaded == true ?
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
