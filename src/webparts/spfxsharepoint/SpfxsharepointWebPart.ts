import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxsharepointWebPart.module.scss';
import * as strings from 'SpfxsharepointWebPartStrings';
import pnp from 'sp-pnp-js';


// import {
//   SPHttpClient
// } from '@microsoft/sp-http';


export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Title: string;
  student_name: string;
  departmennt: string;
  join_date: string;
}
export interface ISpfxsharepointWebPartProps {
  description: string;
}

export default class SpfxsharepointWebPart extends BaseClientSideWebPart<ISpfxsharepointWebPartProps> {

  
  // private _getListData(): Promise<ISPLists> {
  //   return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('student_list')/Items`, SPHttpClient.configurations.v1)
  //     .then((response: any) => {
  //       debugger;
  //       return response.json();
  //     });
  // }
  private _getListData(): Promise<ISPList[]> {
    return pnp.sp.web.lists.getByTitle("student_list").items.get().then((response) => {
     
       return response;
     });
       
    }
  private _renderListAsync(): void {


    this._getListData()
      .then((response) => {
        this._renderList(response);
      });

  }
  private _renderList(items: ISPList[]): void {
    let htm: string = '<div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">Student Details</div>';
    let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';
    html += `<th>Title</th><th>studentName</th><th>Department</th><th>join_date</th>`;
    items.forEach((item: ISPList) => {
      html += `  
         <tr>  
        <td>${item.Title}</td>  
        <td>${item.student_name}</td>  
        <td>${item.departmennt}</td>  
        <td>${item.join_date}</td>  
        </tr>  
        `;
    });
    html += `</table>`;
    const heading: Element = this.domElement.querySelector('#heading');
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    heading.innerHTML = htm;
    listContainer.innerHTML = html;
  }
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.spfxsharepoint}">
    <div class="${ styles.container}">
      <div class="${ styles.row}">
        <div class="${ styles.column}">
          <span class="${ styles.title}">Welcome to SharePoint!</span>
  <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
    <p class="${ styles.description}">${escape(this.properties.description)}</p>
     
          <div id="heading" /> </div>
          <br>
          <div id="spListContainer" />  
            </div>  
          </div>
          </div>
          </div>
          </div>`;
    this._renderListAsync();
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
