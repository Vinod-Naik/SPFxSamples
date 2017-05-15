import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './JsBasicSpWebpart.module.scss';
import * as strings from 'jsBasicSpWebpartStrings';
import { IJsBasicSpWebpartWebPartProps } from './IJsBasicSpWebpartWebPartProps';

//Vinod : Import MockData from Mockdata file
import MockData from './MockData';

//Vinod : Import HttpClient to make Rest calls
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'

//Vinod : Import Environment type to identify the current environment is local or SP Site
import { Environment,EnvironmentType } from '@microsoft/sp-core-library'; 

//Vinod : Interface to define return data types from Rest calls for List and list colllection
export interface ISPLists{
  value : ISPList[]
}

export interface ISPList{
  Title : string;
  Id : string;
}


export default class JsBasicSpWebpartWebPart extends BaseClientSideWebPart<IJsBasicSpWebpartWebPartProps> {
//Vinod : Method to get the list data from Sharepoint using Rest calls
  private _getListData():Promise<ISPLists>{
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$filter=Hidden eq false", SPHttpClient.configurations.v1)
      .then((data: SPHttpClientResponse)=>{
        return data.json();
      }) as Promise<ISPLists>;
  }

  //Vinod : Method to get the data from MockData class for local workbench
  private _getMockData():Promise<ISPLists>{
    return MockData.get(this.context.pageContext.web.absoluteUrl)
      .then((data:ISPList[])=>{
        var listData : ISPLists = {value : data};
        return listData; 
      }) as Promise<ISPLists>;
  }

  //Vinod: Main method to query data
  private _rederListDataAsync() : void{
    //Check environement type and route source
    if(Environment.type == EnvironmentType.Local){
      this._getMockData().then((response)=>{
          this._renderLists(response.value);
      });
    }
    else if(Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint){
        this._getListData().then((response)=>{
            this._renderLists(response.value);
        });
    }
  }

  //Vinod: Method to create the HTML content for the output lists
  private _renderLists(items : ISPList[]):void {
    let html : string = '';
    items.forEach((item: ISPList) => {
      html += `
        <ul class="${styles.list}">
            <li class="${styles.spListItem}">
                <span class="ms-font-l">${item.Title}</span>
            </li>
        </ul>`;
    });
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }
 
 public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">SharePointFramework List Manager!</span>
              <p class="ms-font-l ms-fontColor-white">WebPart Description : ${escape(this.properties.description)}</p>
              <p class="ms-font-l ms-fontColor-white">Current Web : ${escape(this.context.pageContext.web.title)}</p>              
            </div>
          </div>
          <div id="spListContainer" />
        </div>
        
      </div>`;
      this._rederListDataAsync();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
