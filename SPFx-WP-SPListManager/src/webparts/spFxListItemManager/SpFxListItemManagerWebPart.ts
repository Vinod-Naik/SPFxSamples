import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpFxListItemManager.module.scss';
import * as strings from 'spFxListItemManagerStrings';
import { ISpFxListItemManagerWebPartProps } from './ISpFxListItemManagerWebPartProps';

//My references
import { Environment, EnvironmentType} from "@microsoft/sp-core-library";
import { SPHttpClient,SPHttpClientResponse } from "@microsoft/sp-http";
import { ISPList,ISPListItem,ISPLists, ISPListItems } from './ISPDataTypes'
export default class SpFxListItemManagerWebPart extends BaseClientSideWebPart<ISpFxListItemManagerWebPartProps> {

  private _listOptions : IPropertyPaneDropdownOption[] = [];
  private _listDropdownDisabled : boolean = true;
  private _selectedListId : string = "";
  private listItemEntityTypeFullName : string;

  public render(): void {
    this.domElement.innerHTML = `
       <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <span class="ms-font-xl ms-fontColor-dark">SPFx List Manager</span>
            <div class="${styles.row}">
              <span class=ms-font-xl ms-fontColor-dark">${this.properties.targetList}</span>
            </div>
            <div class="${styles.row}">
              <button class="${styles.button} readall-Button">
                <span class="${styles.label}">Read all items</span>
              </button>
              &nbsp;
              <button class="${styles.button} getRecent-Button">
                <span class="${styles.label}">Get recent item</span>
              </button>
            </div>
            <div class="${styles.row}">
              <button class="${styles.button} create-Button">
                <span class="${styles.label}">Create ListItem</span>
              </button>
              &nbsp;
              <button class="${styles.button} update-Button">
                <span class="${styles.label}">Update Item</span>
              </button>
              &nbsp;
              <button class="${styles.button} delete-Button">
                <span class="${styles.label}">Delete Item</span>
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
      this.listItemEntityTypeFullName = undefined;
      this._updateOutputStatus(this.listNotConfigured() ? 'Please configure list in Web Part properties' : 'Ready');
      this._clearOutputDiv();
      this._setButtonHandlers();
      this._enableButtons();

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //Disable reactive mode
  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  //Property Pane Configuration
  protected onPropertyPaneConfigurationStart() : void {
    console.log("onPropertyPaneConfigurationStart called");
    //this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');
    if(this._listOptions.length > 0){
      console.log("Lists already loaded");
      return; 
    }
    this.getListsOptions();
    this.context.propertyPane.refresh()
    
  }

  protected getListsOptions() : void {
    this.getListsAsync()
      .then((response) =>{
        this._listOptions.push({key: "", text: "Select"});
        response.value.map((list : ISPList)=>{
            this._listOptions.push({key: list.Title, text: list.Title});
        })      
      });
  }

  protected getListsAsync() : Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + 
        "/_api/web/lists?$filter=((Hidden eq false) and (BaseTemplate eq 100))&$orderby=Title",
    SPHttpClient.configurations.v1)
      .then((response : SPHttpClientResponse)=>{
        return response.json()
      })
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
                PropertyPaneDropdown('targetList',{
                  label : strings.ListFieldLabel,
                  options : this._listOptions,
                  selectedKey : this.properties.targetList,
                  disabled : false
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private listNotConfigured(): boolean {
    return this.properties.targetList === undefined ||
      this.properties.targetList === null ||
      this.properties.targetList === "" ||
      this.properties.targetList === "Select";
  }

  private _setButtonHandlers():void {
    const wpObj: SpFxListItemManagerWebPart = this;
    this.domElement.querySelector('button.readall-Button').addEventListener('click', ()=> { wpObj._clearOutputDiv(); wpObj.getListItems() });
    this.domElement.querySelector('button.getRecent-Button').addEventListener('click', ()=> { wpObj._clearOutputDiv(); wpObj.getLatestItemId();  });
    this.domElement.querySelector('button.create-Button').addEventListener('click', ()=> { wpObj._clearOutputDiv(); wpObj.createListItem(); });
    this.domElement.querySelector('button.update-Button').addEventListener('click', ()=> { wpObj._clearOutputDiv();  wpObj.updateListItem();});
    this.domElement.querySelector('button.delete-Button').addEventListener('click', ()=> { wpObj._clearOutputDiv(); wpObj.deleteListItem();});
  }

  private _enableButtons():void{
    if(this.listNotConfigured()){
      this.domElement.querySelector('button.readall-Button').setAttribute('disabled','disabled');
      this.domElement.querySelector('button.getRecent-Button').setAttribute('disabled','disabled');
      this.domElement.querySelector('button.create-Button').setAttribute('disabled','disabled');
      this.domElement.querySelector('button.update-Button').setAttribute('disabled','disabled');
      this.domElement.querySelector('button.delete-Button').setAttribute('disabled','disabled');
    }
    else{
      this.domElement.querySelector('button.readall-Button').removeAttribute('disabled');
      this.domElement.querySelector('button.getRecent-Button').removeAttribute('disabled');
      this.domElement.querySelector('button.create-Button').removeAttribute('disabled');
      this.domElement.querySelector('button.update-Button').removeAttribute('disabled');
      this.domElement.querySelector('button.delete-Button').removeAttribute('disabled');
    }
  }

  private _clearOutputDiv():void {
    this.domElement.querySelector('#lstOutput').innerHTML = "";
    this.domElement.querySelector('#lblStatus').innerHTML = "";
  }
  
  private _updateOutputStatus(output : string):void {
    this.domElement.querySelector('#lblStatus').innerHTML = output;
    console.log(output);
  }

  //Handers
  private printListEntityType() : void {
    console.log(this.properties.targetList);
    this.getListItemEntityTypeName()
      .then((response : string) => {
        this._updateOutputStatus(this.properties.targetList + " `n List Entity Type : " + response);
      });
  }

  // All ListItem Operations

  //Not working, returns along with Metadata
  private getListEntityType(): Promise<string> {
    return new Promise<string>((resolve : (listItemEntityTypeFullName : string)=> void, reject : (error ?: any) => void) =>{
      if(this.listItemEntityTypeFullName || this.listItemEntityTypeFullName != undefined){
        resolve(this.listItemEntityTypeFullName)
      }
    
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.properties.targetList}')?$select=ListItemEntityTypeFullName`,        
          SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': "application/json;odata=verbose",
            'odata-version': ''
          }
        })
          .then((response : SPHttpClientResponse) : Promise<{ ListItemEntityTypeFullName : string }> =>{
            return response.json()
          },
          (error : any): void => {
            reject(error);
          })
         .then((response : { ListItemEntityTypeFullName : string }) => {
            this.listItemEntityTypeFullName = response.ListItemEntityTypeFullName,
            console.log(this.listItemEntityTypeFullName);
            resolve(this.listItemEntityTypeFullName);
          })

    });    
  }

  //Working metthos to get EntityType
  private getListItemEntityTypeName(): Promise<string> {
    return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
      if (this.listItemEntityTypeFullName) {
        resolve(this.listItemEntityTypeFullName);
        return;
      }

      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.targetList}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { ListItemEntityTypeFullName: string }): void => {
          this.listItemEntityTypeFullName = response.ListItemEntityTypeFullName;
          resolve(this.listItemEntityTypeFullName);
        });
    });
  }

  private getListItems():void {
    this._updateOutputStatus("Loading list items");
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.properties.targetList}')/items`,
        SPHttpClient.configurations.v1)
        .then((response : SPHttpClientResponse ) : Promise<{value : ISPListItem[]}> => {
            return response.json();
          },
          (error ?:any)=>{
            this._updateOutputStatus("Error occured : " + error);
          })
          .then((response : {value : ISPListItem[]})=>{
              return response.value;
          })
          .then((items : ISPListItem[])=>{
            let html :string = "";
            this._updateOutputStatus("Total Items returned : " + items.length);
            for (let i: number = 0; i < items.length; i++) {
              html += `<li>${items[i].Title} (${items[i].Id})</li>`;
            }
            const listContainer : Element = this.domElement.querySelector('#lstOutput');
            listContainer.innerHTML = html;
          });
  }

  private getLatestItemId(): Promise<number> {
    this._updateOutputStatus("Querying list " + this.properties.targetList);
    return new Promise<number>((resolve : ( itemId : number) => void, reject : (error ?: any)=> void) => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.properties.targetList}')/items?$orderby=Id desc&$top=1`,
          SPHttpClient.configurations.v1)
        .then((response : SPHttpClientResponse) : Promise<{value : {Id : number}[]}>=>{
              return response.json()
            },
            (error ?:any)=>{
              this._updateOutputStatus("Error occured : " + error);
              reject(error);
            })
            .then((response : {value : {Id : number}[]})=>{
                if(response.value.length == 0)
                  resolve(-1)
                else
                  this._updateOutputStatus("Las Item : " + response.value[0].Id);
                  resolve(response.value[0].Id);
            });
    });
  }

  private createListItem():void {
    this.getListItemEntityTypeName()
      .then((listItemEntityTypeFullName : string) : Promise<SPHttpClientResponse> =>{ 
        const curTime : Date = new Date()
        const rndNum : string = `${curTime.getHours()}:${curTime.getMinutes()}:${curTime.getSeconds()}` ;
        const itemTitle :string = `SPFxListItem CreatedOn ${rndNum}`;
        const body:string = JSON.stringify({
          'Title':itemTitle,
          '__metadata':{'type':listItemEntityTypeFullName}
        });

        return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.properties.targetList}')/items`,
          SPHttpClient.configurations.v1,
          {
            headers : {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body : body
          });
        })
        .then((response : SPHttpClientResponse ) => {
            if(response.ok)
              return response.json()
            else
              this._updateOutputStatus(response.status + " : " + response.statusText);
          },
          (error ?:any)=>{
            this._updateOutputStatus("Error occured : " + error);
          })
          .then((listItem : ISPListItem )=>{
              this._updateOutputStatus(`New List Item created : ${listItem.Title}`);
              this.getListItems();
            })
  }

  private updateListItem() : void {
    this._updateOutputStatus("Fetching List Entity Type");
    let itemId2Update : number = undefined;
    let etag : string = undefined;

    this.getListItemEntityTypeName()
    .then((entityName : string ) => {
      this._updateOutputStatus("Fetching Itemid to Update");
      return this.getLatestItemId();
    })
    .then((itemId : number) => { 
      if(itemId == -1){
        this._updateOutputStatus("Item not found");
        return;
      }
      itemId2Update = itemId;
      this._updateOutputStatus("Loading latest item to update")
      return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.properties.targetList}')/items(${itemId})?$select=Id`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
    })
    .then((response : SPHttpClientResponse ) : Promise<ISPListItem> =>  {
        etag = response.headers.get('ETag');
        return response.json();
    })
    .then((item : ISPListItem ) : Promise<SPHttpClientResponse> =>  {
        const curTime : Date = new Date()
        const rndNum : string = `${curTime.getHours()}:${curTime.getMinutes()}:${curTime.getSeconds()}` ;
        const itemTitle :string = `SPFxListItem UpdatedOn ${rndNum}`;
        const body:string = JSON.stringify({
          'Title':itemTitle,
          '__metadata':{'type':this.listItemEntityTypeFullName }
        });
        this._updateOutputStatus("Updating item with itemid: " + item.Id);
        return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.targetList}')/items(${item.Id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': '',
              'IF-MATCH': etag,
              'X-HTTP-Method': 'MERGE'
            },
            body: body
          });
    })
    .then((response : SPHttpClientResponse ) => {
        if(response.ok){
          this._updateOutputStatus("Updated item with ID : " + itemId2Update);
          this.getListItems();
        }
        else{
          this._updateOutputStatus("Error occured : " + response.status + " : " + response.statusText);
        }
    },
    (error : any)=>{
      this._updateOutputStatus("Error occured : " + error);
    });

  }

  private deleteListItem() : void {
     this._updateOutputStatus("Fetching List item to delete");
     let itemId2Delete : number = undefined;
     let etag : string = undefined;
     this.getLatestItemId()
      .then((itemId : number ) : Promise<SPHttpClientResponse>=>{ 
        return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.targetList}')/items(${itemId})`,
          SPHttpClient.configurations.v1,
          {
            headers : {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      })
      .then((response : SPHttpClientResponse ) : Promise<ISPListItem> =>  {
          etag = response.headers.get('ETag');
          return response.json();
      })
      .then((item : ISPListItem ) : Promise<SPHttpClientResponse> =>  {
        return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.targetList}')/items(${item.Id})`,
            SPHttpClient.configurations.v1,
            {
              headers : {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': '',
                'IF-MATCH': etag,
                'X-HTTP-Method': 'DELETE'
              }
            });
      })
      .then((response : SPHttpClientResponse ) => {
          if(response.ok){
            this._updateOutputStatus("Deleted item with ID : " + itemId2Delete);
            this.getListItems();
          }
          else{
            this._updateOutputStatus("Error occured : " + response.status + " : " + response.statusText);
          }
      },
      (error : any)=>{
        this._updateOutputStatus("Error occured : " + error);
      });
  }
  
}
