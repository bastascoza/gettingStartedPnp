import { Version, Log } from "@microsoft/sp-core-library";
const LOG_SOURCE: string = "GloboWebPart";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

import styles from "./GloboWebPart.module.scss";
import * as strings from "GloboWebPartStrings";

export interface IGloboWebPartProps {
  description: string;
  showStaffNumber: Boolean;
}

export default class GloboWebPart extends BaseClientSideWebPart<IGloboWebPartProps> {
  public render(): void {
    //  Log.verbose(LOG_SOURCE,'Hello from Globomantics web part',this.context.serviceScope)
    this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        //get information about the current user from the microsoft Graph
        client.api("/me").get((error, userProfile: any, rawResponse?: any) => {
          this.domElement.innerHTML = `
      <div class="${styles.globo}">
        <div class="${styles.container}">
          <div class="${styles.row}">            
              <span class="${styles.title}">Welcome ${escape(
            this.context.pageContext.user.displayName
          )}!</span>
              <div class="${styles.subTitle}" id="spUserContainer"></div>
              <div class="${styles.rowTable}">
                    <div class="${styles.columnTable3}">
                        <h2> Manager </h2>
                        <div id="spManager"></div>
                    </div>                  
                    <div class="${styles.columnTable3}">
                        <h2> Colleagues </h2>
                        <div id="spColleagues"></div>
                    </div>                  
                    <div class="${styles.columnTable3}">
                        <h2> Direct Reports </h2>
                        <div id="spReports"></div>
                    </div>                  
              </div>
          </div>
        </div>
      </div>`;
          this._renderJobtitle(userProfile);
          if (this.properties.showStaffNumber) {
            this._renderEmployeeId(client)
          }
          this._renderDirecReports(client);
        });
      });
  }
  private _renderColleagues(client:MSGraphClient, userProfile:MicrosoftGraph.User, managerId:string):void {
    client.
    api(`/users/${managerId}/directReports`)
    .get((error,directReports:any,rawResponse?:any) => {
      const spColleageContainer: Element = this.domElement.querySelector('#spColleagues');
      let html:string=spColleageContainer.innerHTML;
      directReports.value.forEach((directReport:MicrosoftGraph.User)=>{
        if(directReport.id!=userProfile.id){
          html+=`<p class="${styles.description}">${escape(directReport.displayName)}</p>`;
        }
      });
      spColleageContainer.innerHTML=html;
    });
  }
  private _RenderManagerAndColleagues(client:MSGraphClient, userProfile: MicrosoftGraph.User):void{
    client.api('/me/manager')
    .get((error,manager:any,rawResponse?:any)=>{
      const spContainer:Element=this.domElement.querySelector('#spManager')
      let html:string = spContainer.innerHTML;
      if(manager!=null){
         html+=`<p class="${styles.description}"> ${escape(manager.displayName)}</p>`;
         spContainer.innerHTML=html;
      }
     this._renderColleagues(client,userProfile,manager.id)      
    })
  }

  private _renderDirecReports(client:MSGraphClient):void{
    client.api('/me/directReports')
    .get((error,directReports:any,rawResponse?:any)=>{
      const spContainer:Element=this.domElement.querySelector('#spReports')
      let html:string = spContainer.innerHTML;
      directReports.value.forEach((directReport:MicrosoftGraph.User) => {
        html+=`<p class="${styles.description}"> ${escape(directReport.displayName)}</p>`
      });
      spContainer.innerHTML=html;
    })
  }
  private _renderJobtitle(userProfile: MicrosoftGraph.User): void {
    const spUserContainer: Element =
      this.domElement.querySelector("#spUserContainer");
    let html: string = spUserContainer.innerHTML;
    html += `<p>${escape(userProfile.jobTitle)}</p>`;
    spUserContainer.innerHTML = html;
  }
  private _renderEmployeeId(client: MSGraphClient): void {
    client
      .api("/me/employeeId/$value")
      .get((error, employeeId: any, rawResponse?: any) => {
        const spUserContainer: Element =
          this.domElement.querySelector("spUserContainer");
        spUserContainer.innerHTML += `<p>${escape(employeeId)}</p>`;
      });
  }
  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyPaneToggle("ShowStaffNumber", {
                  label: "Show Staff Number",
                  checked: true,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
