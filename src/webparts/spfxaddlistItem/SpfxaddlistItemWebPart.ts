import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxaddlistItem.module.scss';
import * as strings from 'spfxaddlistItemStrings';
import { ISpfxaddlistItemWebPartProps } from './ISpfxaddlistItemWebPartProps';

export default class SpfxaddlistItemWebPart extends BaseClientSideWebPart<ISpfxaddlistItemWebPartProps> {

  public render(): void {
     this.domElement.innerHTML = `
      <div class="${styles.listItem}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col" style="width:100%">
              <span><h1>SharePoint List Item Manipulation</h1></span>
            </div>
            <div class="ms-Grid-col" style="width:100%">
              <span>Title : </span>
            </div>
            <div class="ms-Grid-col" style="width:100%">
              <input type="text" id="Title" style="width:100%">
            </div>
            <div class="ms-Grid-col" style="width:100%">
              <span>Description : </span>
            </div>
            <div class="ms-Grid-col" style="width:100%">
              <textarea type="text" id="description" rows=5 style="width:100%"></textarea>
            </div>
            <div class="ms-Grid-col" style="width:100%;text-align:center">
              <button type="button" id="btn_add" style="background-color:#009688;color:white;font-weight:bold">Create New ListItem</button>
            </div>
          </div>

        </div>
        <div class="status" style="color:red"></div>
      </div>
      <div >
      `;

      const events: SpfxaddlistItemWebPart = this;
      var button = document.querySelector('#btn_add');
      button.addEventListener('click', () => { events.CreateNewItem(); });
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
   private CreateNewItem(): void {
    this.usermessage('Creating list Item ...');
    let title = (<HTMLInputElement>document.getElementById("Title")).value.trim();
    let description = (<HTMLInputElement>document.getElementById("description")).value.trim();
    if(title != '' && description != '')
    {
      //Create a array object with all column values
        let requestdata = {};
        requestdata['Title'] = title;
        requestdata['Description'] = description;
        this.usermessage('Creating list Item ...' + requestdata);
        this.addListItem('spfxlist',requestdata);
    }
    else
    {
        if(title == '' && description == '')
        {this.usermessage('Please enter title and description');}
        else if(title == '')
        {this.usermessage('Please enter title');}
        else if(description == '')
        {this.usermessage('Please enter description');}
    }
  }

  private addListItem(listname:string,requestdata:{})
  {   
      let requestdatastr = JSON.stringify(requestdata);
      requestdatastr = requestdatastr.substring(1, requestdatastr .length-1);
      console.log(requestdatastr);
      let requestlistItem: string = JSON.stringify({
        '__metadata': {'type': this.getListItemType(listname)}
      });
      requestlistItem = requestlistItem.substring(1, requestlistItem .length-1);
      requestlistItem = '{' + requestlistItem + ',' + requestdatastr + '}';
      console.log(requestlistItem);
      this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listname}')/items`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: requestlistItem
          })
          .then((response: SPHttpClientResponse): Promise<IListItem> => {
            console.log('response.json()');
          return response.json();
        })
         .then((item: IListItem): void => {
            console.log('Creation');
            this.usermessage(`List Item created successfully... '(Item Id: ${item.Id})`);
      }, (error: any): void => {
          this.usermessage('List Item Creation Error...');
        });      
  }

  private usermessage(status: string): void {
    this.domElement.querySelector('.status').innerHTML = status;
  }
  private getListItemType(name: string) {
	  let safeListType = "SP.Data." + name[0].toUpperCase() + name.substring(1) + "ListItem";
	  safeListType = safeListType.replace(/_/g,"_x005f_");
	  safeListType = safeListType.replace(/ /g,"_x0020_");
    return safeListType;
}
}

export interface IListItem {
  Title: string;
  Description: string;
  Id: number;
}                                                                                                                                                                                                                                                                                                                                                                                                                                   