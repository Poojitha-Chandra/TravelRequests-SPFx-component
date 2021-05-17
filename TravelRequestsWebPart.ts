import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './TravelRequestsWebPart.module.scss';
import * as strings from 'TravelRequestsWebPartStrings';
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
export interface IReadSpListItemsWebPartProps
 {
  description: string;
}
export interface ISPLists 
{
  value: ISPList[];
}

export interface ISPList 
{
  Title: string;
  Place: string;
  TravelDescription: string;
  DateofTravel : string;
  DateofReturn: string;
  ProjectName: string;
  ModeOfJourney: string;
}
export interface ITravelRequestsWebPartProps {
  description: string;
}

export default class TravelRequestsWebPart extends BaseClientSideWebPart<ITravelRequestsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.travelRequests }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
          <h2>Travel Request</h2>
          <h2>Request Form</h2>
          <hr/>
              <div class="${ styles.column}">
              Name:
              </div>
              <div class="${ styles.column}">
              <input type='text' id='txtName'/>
              </div>
          <br/>
              <div class="${ styles.column}">
              Place:
              </div>
              <div class="${ styles.column}">
              <input type='text' id='txtPlace'/>
              </select>          
              </div>
            <br/>
              <div class="${ styles.column}">
              Travel Description:
              </div>
              <div class="${ styles.column}">
              <input type='text' id='txtTravelDescription'/>
              </div>
              <br/>
              <div class="${ styles.column}">
              Date of Travel:
              </div>
              <div class="${ styles.column}">
              <input type='date' id='txtDate'/>
              </div>
              <br/>
              <div class="${ styles.column}">
              Date of Return:
              </div>
              <div class="${ styles.column}">
              <input type='date' id='txtDateRet'/>
              </div>
              <br/>
              <div class="${ styles.column}">
              Project Name:
              </div>
              <div class="${ styles.column}">
              <input type='text' id='txtProjectName'/>
              </div>
              <br/>
              <div class="${ styles.column}">
              Mode Of Journey:
              </div>
              <div class="${ styles.column}">
              <select id="ddlModeOfJourney">
              <option value="Train">Train</option>
              <option value="Flight">Flight</option>
              <option value="Car">Car</option>
              <option value="Others">Others</option>
              </select>          
              </div>
              <br/>
              <div class="${ styles.column}">
              <input type="submit" value="Create request" id="btnSubmit"><br/>
              <div id="spListCreateRequest"/>
              </div>
            </div>
          </div>
        <div class="${ styles.container}">  
        <div class="${ styles.row}">
          <hr/>
          <h2>Delete Travel Request</h2>
          <hr/>
          <div class="${ styles.column}">
           Enter Item ID: <input type='text' id='txtItemIDToDelete'/> <input type="submit" value="Delete List Item" id="btnDelete">
         <br/>
           <div id="spListItemDeleteStatus" />
          </div>
          </div>
          </div>
            </div>`;
            this._setButtonEventHandlers();
        }
   private _setButtonEventHandlers(): void {
          //this.domElement.querySelector("#btnSubmit").addEventListener("click", () => { this._renderListAsync(); });
      
          this.domElement.querySelector('#btnSubmit').addEventListener('click', () => { this._renderListAsync(); });
          this.domElement.querySelector('#btnDelete').addEventListener('click', () => { this._deleteListItemByID(); });
      }
      private _getListData(): Promise<ISPLists> {
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/Travel/_api/web/lists/GetByTitle('TravelRequests')/Items",SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
              console.log(response.status);
            return response.json();
            }
        );
      }
      private _renderListAsync(): void {
        if (Environment.type == EnvironmentType.SharePoint || 
                 Environment.type == EnvironmentType.ClassicSharePoint) {
         this._getListData()
           .then((response) => {
             this._renderList(response.value);
           });
       }
      }
      private _renderList(items: ISPList[]): void {
        var title = document.getElementById("txtName")["value"];
        var travelDate = document.getElementById("txtDate")["value"];
              //console.log(travelDate);
        let flag = 0;
        items.forEach((item: ISPList) => {
          if(item.Title == title){
            //console.log(item.TravelDate);
            //const date = item.TravelDate.getFullYear() + item.TravelDate.getMonth() + item.TravelDate.getDate();
            //console.log(date);
            console.log(item.DateofTravel.substring(0,10));
            if(item.DateofTravel.substring(0,10) == travelDate){
              alert("employee already has a travel request with the same date");
              //return false;
              flag = 1;
            }
          }
        });
        //return true;
        if(flag == 0){
          this.createTravelRequest();
        }
       }
      private createTravelRequest(): void {
        var title = document.getElementById("txtName")["value"];
        var place = document.getElementById("txtPlace")["value"];
        var travelDescription = document.getElementById("txtTravelDescription")["value"];
        var dateOfTravel = document.getElementById("txtDate")["value"];
        var dateOfReturn = document.getElementById("txtDateRet")["value"];
        var projectName = document.getElementById("txtProjectName")["value"];
        var modeOfJourney = document.getElementById("ddlModeOfJourney")["value"];
      
        const url: string = this.context.pageContext.site.absoluteUrl +"/Travel"+ "/_api/web/lists/getbytitle('TravelRequests')/items";
        const itemDefinition: any = {
          "Title": title,
          "Place": place,
          "TravelDescription": travelDescription,
          "DateofTravel": dateOfTravel,
          "DateofReturn": dateOfReturn,
          "ProjectName": projectName,
          "ModeOfJourney": modeOfJourney
      
        };
        const spHttpClientOptions: ISPHttpClientOptions = {
          "body": JSON.stringify(itemDefinition)
        };
      
        this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
          .then((response: SPHttpClientResponse) => {
            if (response.status === 201) {
              let message: Element = this.domElement.querySelector('#spListCreateRequest');
              message.innerHTML = "Create: List Item created successfully.";
              this.clear();
            } else {
              let message: Element = this.domElement.querySelector('#spListCreateRequest');
              message.innerHTML = "Create: List Item creation failed. " + response.status + " - " + response.statusText;
            }
          });
      }
      
      private clear(): void {
        document.getElementById("txtName")["value"] = '';
        document.getElementById("txtPlace")["value"] = '';
        document.getElementById("txtTravelDescription")["value"] = '';
        document.getElementById("txtDate")["value"] = '';
        document.getElementById("txtDateRet")["value"] = '';
        document.getElementById("txtProjectName")["value"] = '';
        document.getElementById("ddlModeOfJourney")["value"] = 'Train';
      }
      private _deleteListItemByID(): void {
        let id: string = document.getElementById("txtItemIDToDelete")["value"];
        const url: string = this.context.pageContext.site.absoluteUrl + "/Travel"+"/_api/web/lists/getbytitle('TravelRequests')/items(" + id + ")";
      
        const headers: any = { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" };
        const spHttpClientOptions: ISPHttpClientOptions = {
          "headers": headers
        };
        this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
          .then((response: SPHttpClientResponse) => {
            if (response.status === 204) {
              let message: Element = this.domElement.querySelector('#spListItemDeleteStatus');
              message.innerHTML = "Delete: List Item deleted successfully.";
             this.clearDelete();
            } else {
              let message: Element = this.domElement.querySelector('#spListItemDeleteStatus');
              message.innerHTML = "List Item delete failed." + response.status + " - " + response.statusText;
            }
          });
      }
   private clearDelete(): void {
        document.getElementById("txtItemIDToDelete")["value"] = "";
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
