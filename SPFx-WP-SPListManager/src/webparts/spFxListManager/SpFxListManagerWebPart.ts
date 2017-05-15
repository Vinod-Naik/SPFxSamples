import { Version } from '@microsoft/sp-core-library';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpFxListManager.module.scss';
import * as strings from 'spFxListManagerStrings';
import { ISpFxListManagerWebPartProps } from './ISpFxListManagerWebPartProps';

// My Updated/Added Code Here

import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {ISPList, ISPLists } from './ISPDataTypes'
import MockData from './MockData'




export default class SpFxListManagerWebPart extends BaseClientSideWebPart<ISpFxListManagerWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <span class="ms-font-xl ms-fontColor-dark">SPFx List Manager</span>
            <div class="${styles.row}">
              <button class="${styles.button} readall-Button">
                <span class="${styles.label}">Read all lists</span>
              </button>
              &nbsp;
              <button class="${styles.button} getRecent-Button">
                <span class="${styles.label}">Get List</span>
              </button>
            </div>
            <div class="${styles.row}">
              <button class="${styles.button} create-Button">
                <span class="${styles.label}">Create List</span>
              </button>
              &nbsp;
              <button class="${styles.button} update-Button">
                <span class="${styles.label}">Update list</span>
              </button>
              &nbsp;
              <button class="${styles.button} delete-Button">
                <span class="${styles.label}">Delete List</span>
              </button>
            </div>
            <div class="${styles.row}">
              <ul id="lstOutput">              
              </ul>
            </div>
            <div class="${styles.row}">              
              <span class="${styles.label}" id="lblStatus"></span>
            </div>
          </div>
        </div>
      </div>`;
      this._enableButtons();
      this._clearOutputDiv();
      this._setButtonHandlers()
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

  // My Custom Code goes here

  private _setButtonHandlers():void {
    const wpObj: SpFxListManagerWebPart = this;
    this.domElement.querySelector('button.readall-Button').addEventListener('click', ()=> { wpObj._clearOutputDiv(); wpObj._renderListDataAsync(); });
    this.domElement.querySelector('button.getRecent-Button').addEventListener('click', ()=> { wpObj._clearOutputDiv(); wpObj._getRecentListTitle(); });
    this.domElement.querySelector('button.create-Button').addEventListener('click', ()=> { wpObj._clearOutputDiv(); wpObj._createListAync(); });
    this.domElement.querySelector('button.update-Button').addEventListener('click', ()=> { wpObj._clearOutputDiv(); wpObj._updateListAsync(); });
    this.domElement.querySelector('button.delete-Button').addEventListener('click', ()=> { wpObj._clearOutputDiv(); wpObj._deleteList(); });
  }
  private _enableButtons():void{
    if(this._isLocal()){
      this.domElement.querySelector('button.create-Button').setAttribute('disabled','disabled');
      this.domElement.querySelector('button.update-Button').setAttribute('disabled','disabled');
      this.domElement.querySelector('button.delete-Button').setAttribute('disabled','disabled');
    }
    else{
      this.domElement.querySelector('button.create-Button').removeAttribute('disabled');
      this.domElement.querySelector('button.update-Button').removeAttribute('disabled');
      this.domElement.querySelector('button.delete-Button').removeAttribute('disabled');
    }
  }
  // Button Handlers

  private _clearOutputDiv():void {
    this.domElement.querySelector('#lstOutput').innerHTML = "";
    this.domElement.querySelector('#lblStatus').innerHTML = "";
  }

  private _updateOutputStatus(output : string):void {
    this.domElement.querySelector('#lblStatus').innerHTML = output;
    console.log(output);
  }

  private _isLocal():boolean{
    if((Environment.type == EnvironmentType.ClassicSharePoint) || (Environment.type == EnvironmentType.SharePoint)){
      console.log("Running on Sharepoint");
      return false;
    }
    else if(Environment.type == EnvironmentType.Local){
      console.log("Running locally");
      return true;
    }
  }
 //Render All Lists
  private _renderListDataAsync():void{
    if(this._isLocal()){
      this._getMockData().then((response)=>{
          this._renderLists(response.value);
      });
    }
    else{
        this._getSPLists().then((response)=>{
            this._renderLists(response.value);
        });
    }
  }
  
  private _renderLists(items : ISPList[]):void{
    let html :string = "";
    items.forEach((item : ISPList)=> {
      html += `<li> ${item.Title} </li>`
    });
    const listContainer : Element = this.domElement.querySelector('#lstOutput');
    listContainer.innerHTML = html;
  }

  private _getSPLists(): Promise<ISPLists>{
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + 
          "/_api/web/lists?$filter=Hidden eq false", 
          SPHttpClient.configurations.v1)
        .then((response : SPHttpClientResponse)=>{
          return response.json();
        });
  }
  private _getMockData():Promise<ISPLists>{
    return MockData.get(this.context.pageContext.web.absoluteUrl)
      .then((data:ISPList[])=>{
        var listData : ISPLists = {value : data};
        return listData; 
      }) as Promise<ISPLists>;
  }

  // Create List

  private _createListAync():void {
    const rndNum : string = (new Date().getHours().toString()) + (new Date().getMinutes().toString()) ;
    const listTitle :string = `SPFxList ${rndNum}`;
    const desc: string = `Created on ${new Date().toUTCString()}`;
    const body:string = JSON.stringify({
      'Title':listTitle,
      'BaseTemplate': 100,
      'Description': desc,
      '__metadata':{'type':'SP.List'}
    });
    
    this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl + "/_api/web/lists",
          SPHttpClient.configurations.v1,
          {
            // For post operations additional paramter with headers and body needs to sent
            headers:{
                  'Accept': 'application/json;odata=nometadata',
                  'Content-type': 'application/json;odata=verbose',
                  'odata-version': ''
                },
                body : body
          })
          .then((response:SPHttpClientResponse) =>{
            // use Nested promises to wait for async operations
            if(response.ok){
              return response.json();
            }
            else{
              return Promise.reject(response.status + " : " + response.statusText);
            }
          })
          .then((listObj : ISPList):void =>{
              this._updateOutputStatus(`New List ${listObj.Title} created with description ${listObj.Description}`);
              this._renderListDataAsync();
          },
          (error:any):void=>{
              this._updateOutputStatus(`Error Occured: ${error}`);
          });

  }

  private _updateListAsync(): void{
    let etag : string = undefined;
    this._updateOutputStatus("Updating List Description ");
    this._getRecentList()
      .then((listName : string)=>{
          this._updateOutputStatus("Getting current list etag");
          this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/lists/getByTitle('${listName}')`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'odata-version': ''
              }
            })
            .then((response : SPHttpClientResponse ) : Promise<ISPList> =>{
              etag = response.headers.get('Etag');
              this._updateOutputStatus("Etag for list found : " + etag );
              return response.json();
            })
            .then((list : ISPList): Promise<SPHttpClientResponse> => {
              this._updateOutputStatus("Updating List " + list.Title);
              const body : string = JSON.stringify({
                  '__metadata':{'type':'SP.List'},
                  'Description': `Updated on ${new Date().toUTCString()}`
              });

              return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/lists/getByTitle('${listName}')`,
                  SPHttpClient.configurations.v1,
                  {
                    headers: {
                      'Accept': 'application/json;odata=nometadata',
                      'Content-type': 'application/json;odata=verbose',
                      'odata-version': '',
                      'IF-MATCH': etag,
                      'X-HTTP-Method': 'MERGE'
                    },
                    body : body
                  });
            })            
              .then((response : SPHttpClientResponse): void => {
                  this._updateOutputStatus("List Updattion successful" + listName);
              },
              (error ?: any)=> {
                  this._updateOutputStatus("List Updation failed : " + error); 
              });
          
      },
      (error ?: any)=>{ 
        this._updateOutputStatus("Error occured" + error)
      });
  }

  private _deleteList() : void {
    let etag : string = undefined;
    this._getRecentList()
      .then((listName : string)=>{
          this._updateOutputStatus("Getting current list etag");
          this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/lists/getByTitle('${listName}')`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'odata-version': ''
              }
            })
            .then((response : SPHttpClientResponse ) : Promise<ISPList> =>{
              etag = response.headers.get('Etag');
              this._updateOutputStatus("Etag for list found : " + etag );
              return response.json();
            })
            .then((list : ISPList): Promise<SPHttpClientResponse> => {
              this._updateOutputStatus("Deleting List " + list.Title);

              return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/lists/getByTitle('${listName}')`,
                  SPHttpClient.configurations.v1,
                  {
                    headers: {
                      'Accept': 'application/json;odata=nometadata',
                      'Content-type': 'application/json;odata=verbose',
                      'odata-version': '',
                      'IF-MATCH': etag,
                      'X-HTTP-Method': 'DELETE'
                    }
                  });
            })            
              .then((response : SPHttpClientResponse): void => {
                  if(response.ok){
                    this._updateOutputStatus(listName + "List deleted successfully : ");
                    this._renderListDataAsync();
                  }
                  else {
                    this._updateOutputStatus("List Deletion failed : " + response.status + " : " + response.statusText); 
                  }                  
              },
              (error ?: any)=> {
                  this._updateOutputStatus("List Deletion failed : " + error); 
              });
          
      },
      (error ?: any)=>{ 
        this._updateOutputStatus("Error occured" + error)
      });
  }
  private _getRecentList():Promise<string>{
    return new Promise<string>((resolve : (listName : string)=> void, reject: (error ?: any) => void) => {
      const baseUrl : string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$filter=Hidden%20eq%20false&$orderby=Created%20desc&$select=Title&$top=1"
      this.context.spHttpClient.get(baseUrl, 
          SPHttpClient.configurations.v1)
        .then((response : SPHttpClientResponse) : Promise<{ value: { Title: string}[] }> =>{
            
            if(response.ok){
                return response.json();
            }
            else {
                reject(response.status + " : " + response.statusText);
            } 
          },
          (error:any):void =>{
              reject(error);
        })
          .then((response : {value : {Title :string}[]}) => {
            if(response.value.length == 0)
              resolve("")
            else
              resolve(response.value[0].Title)
          });
    });

  }

private _getRecentListTitle(): void {
    //const baseUrl : string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$filter=Hidden eq false&$orderby=Created desc&$top=1"
    const baseUrl : string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$filter=Hidden%20eq%20false&$orderby=Created%20desc&$select=Title,Id,Description&$top=1"
    this.context.spHttpClient.get(baseUrl, 
          SPHttpClient.configurations.v1)
        .then((response : SPHttpClientResponse) : Promise<{ value: { Title: string, Description : string }[] }> =>{
          return response.json();
        },
        (error:any):void =>{
            this._updateOutputStatus(`Error Occured: ${error}`);
        }) 
        .then((response : { value: { Title: string, Description : string }[] }) => {
          //Define runtime input types to avoid creation of new types
          if(response.value.length == 0){
            this._updateOutputStatus(`No list Found`);
          }
          else{
            this._updateOutputStatus("Last list is : " + response.value[0].Title + " Desc: " + response.value[0].Description);
          }
        });
  }

}
