import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IListItem } from './IListItem';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NoFrameworkCrudWebPart.module.scss';
import * as strings from 'NoFrameworkCrudWebPartStrings';
import pnp, {sp, Item} from "sp-pnp-js";
export interface INoFrameworkCrudWebPartProps {
  listName: string;
}

export interface ISPListCustomers{
  value:IListItem[];
  }

export default class NoFrameworkCrudWebPart extends BaseClientSideWebPart<INoFrameworkCrudWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.noFrameworkCrud }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">CRUD Operation for SharePoint Framework!</span>
              <p class="${ styles.subTitle }">No Framework</p>
              <p class="${ styles.description }">Name: ${escape(this.properties.listName)}</p>

              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <button class="${styles.button} create-Button">
                    <span class="${styles.label}">Create item</span>
                  </button>
                  <button class="${styles.button} read-Button">
                    <span class="${styles.label}">Read item</span>
                  </button>
                 
                </div>
              </div>

              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                <button class="${styles.button} read">
                <span class="${styles.label}">Read item PNP</span>
              </button>
              <button class="${styles.button} add-Button">
              <span class="${styles.label}">Add item</span>
            </button>
                </div>
                </div>

              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <button class="${styles.button} update-Button">
                    <span class="${styles.label}">Update item</span>
                  </button>
                  <button class="${styles.button} delete-Button">
                    <span class="${styles.label}">Delete item</span>
                  </button>
                </div>
              </div>
              <div>
              <input id="EmployeeName"  placeholder="EmployeeName"/>  
              <input id="EmployeeAddress"  placeholder="EmployeeAddress"/>  
              </div>
              <div id="spListContainer"/>
              </div>
              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <div class="status"></div>
                  <ul class="items"><ul>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>`;
      this.setButtonsEventHandlers();
  }

  private setButtonsEventHandlers(): void {  
    const webPart: NoFrameworkCrudWebPart = this;  
    this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.createItem(); });  
    this.domElement.querySelector('button.add-Button').addEventListener('click', () => { webPart.AddItem(); });  
    this.domElement.querySelector('button.read-Button').addEventListener('click', () => { webPart.readItem(); }); 
    this.domElement.querySelector('button.read').addEventListener('click', () => { webPart.readItemPNP(); }); 
    this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.updateItem(); });  
    this.domElement.querySelector('button.delete-Button').addEventListener('click', () => { webPart.deleteItem(); });  
  } 
  
private _getListCustomerPnp():Promise<IListItem[]>
{
  pnp.setup({
    spfxContext: this.context
  });
return pnp.sp.web.lists.getByTitle(`Employee`).items.
top(100).orderBy("Title").select("Title","EmployeeID",
"EmployeeName","EmployeeAddress","EmployeeType","Author/Id","Author/Title").expand
("Author").get().then
(
(response:any[])=>{
return response;
});
}

private _renderListCustomer(items:IListItem[]):void
{
let html:string=`<table width='100%' border=1>`;
html+=`<thead><th>ID</th><th>Name</th><th>Address</th><th>Type</th><th>Author</th>
`+
`</thead><tbody>`;

items.forEach((item:IListItem)=>
{
html+= `<tr><td>${item.EmployeeID}</td>
<td>${item.EmployeeName}</td>
<td>${item.EmployeeAddress}</td>
<td>${item.EmployeeType}</td>
<td>${item.Author.Title}</td>

</tr>`;
});
html+=`</tbody></table>`;
const listContainer:Element=this.domElement.querySelector("#spListContainer");
listContainer.innerHTML=html;
}
private AddItem(): void {
  pnp.setup({
    spfxContext: this.context
  });
  pnp.sp.web.lists.getByTitle('Employee').items.add({      
    EmployeeName : document.getElementById('EmployeeName')["value"],  
    EmployeeAddress : document.getElementById('EmployeeAddress')["value"]
   
 }); 

  alert("Record with Employee Name : "+ document.getElementById('EmployeeName')["value"] + " Added !"); 
  this.readItemPNP();
}
  
  private createItem(): void {
    const body: string = JSON.stringify({
      'Title': `Item ${new Date()}`
    });

    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: body
    })
    .then((response: SPHttpClientResponse): Promise<IListItem> => {
      return response.json();
    })
    .then((item: IListItem): void => {
      this.updateStatus(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
    }, (error: any): void => {
      this.updateStatus('Error while creating the item: ' + error);
    });
  }
  private readItemPNP(): void {
    this._getListCustomerPnp().then((response)=>
    {
    this._renderListCustomer(response);
    });
  } 
  private readItem(): void {
    this.getLatestItemId()
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        this.updateStatus(`Loading information about item ID: ${itemId}...`);
        
        return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${itemId})?$select=Title,Id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        return response.json();
      })
      .then((item: IListItem): void => {
        this.updateStatus(`Item ID: ${item.Id}, Title: ${item.Title}`);
      }, (error: any): void => {
        this.updateStatus('Loading latest item failed with error: ' + error);
      });
  } 
  
  private updateItem(): void {
    let latestItemId: number = undefined;
    this.updateStatus('Loading latest item...');

    this.getLatestItemId()
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this.updateStatus(`Loading information about item ID: ${itemId}...`);
        
        return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${latestItemId})?$select=Title,Id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        return response.json();
      })
      .then((item: IListItem): void => {
        this.updateStatus(`Item ID1: ${item.Id}, Title: ${item.Title}`);

        const body: string = JSON.stringify({
          'Title': `Updated Item ${new Date()}`
        });

        this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${item.Id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'odata-version': '',
              'IF-MATCH': '*',
              'X-HTTP-Method': 'MERGE'
            },
            body: body
          })
          .then((response: SPHttpClientResponse): void => {
            this.updateStatus(`Item with ID: ${latestItemId} successfully updated`);
          }, (error: any): void => {
            this.updateStatus(`Error updating item: ${error}`);
          });
      });
  } 
  
  private deleteItem(): void {
    if (!window.confirm('Are you sure you want to delete the latest item?')) {
      return;
    }

    this.updateStatus('Loading latest items...');
    let latestItemId: number = undefined;
    let etag: string = undefined;
    this.getLatestItemId()
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this.updateStatus(`Loading information about item ID: ${latestItemId}...`);
        return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${latestItemId})?$select=Id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        etag = response.headers.get('ETag');
        return response.json();
      })
      .then((item: IListItem): Promise<SPHttpClientResponse> => {
        this.updateStatus(`Deleting item with ID: ${latestItemId}...`);
        return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${item.Id})`,
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
      .then((response: SPHttpClientResponse): void => {
        this.updateStatus(`Item with ID: ${latestItemId} successfully deleted`);
      }, (error: any): void => {
        this.updateStatus(`Error deleting item: ${error}`);
      });
  }
  
  private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items?$orderby=Id desc&$top=1&$select=id`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { value: { Id: number }[] }): void => {
          if (response.value.length === 0) {
            resolve(-1);
          }
          else {
            resolve(response.value[0].Id);
          }
        });
    });
  }
  private updateStatus(status: string, items: IListItem[] = []): void {
    this.domElement.querySelector('.status').innerHTML = status;
    this.updateItemsHtml(items);
  }
  private updateItemsHtml(items: IListItem[]): void {
    this.domElement.querySelector('.items').innerHTML = items.map(item => `<li>${item.Title} (${item.Id})</li>`).join("");
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
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
