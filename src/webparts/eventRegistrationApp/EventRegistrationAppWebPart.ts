import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import type { IReadonlyTheme } from '@microsoft/sp-component-base';
//import { escape } from '@microsoft/sp-lodash-subset';
import { IEventRegistration } from './IEventRegistration';
import styles from './EventRegistrationAppWebPart.module.scss';
import * as strings from 'EventRegistrationAppWebPartStrings';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

export interface IEventRegistrationAppWebPartProps {
  description: string;
}

export default class EventRegistrationAppWebPart extends BaseClientSideWebPart<IEventRegistrationAppWebPartProps> {

  // private _isDarkTheme: boolean = false;
  // private _environmentMessage: string = '';

  // functions start
  private CreateItem(): void{

    // get inputs
    const userName = document.getElementById("txtUserName") as HTMLInputElement;
    const userNameValue: string = userName.value;
    const email = document.getElementById("txtEmail") as HTMLInputElement;
    const emailValue: string = email.value;
    const batch = document.getElementById("ddlBatch") as HTMLSelectElement;
    const batchValue: string = batch.value;
    const levelOfKnowledge = document.getElementById("ddlLevelOfKnowledge") as HTMLSelectElement;
    const levelOfKnowledgeValue: string = levelOfKnowledge.value;
    //const siteURL: string  = this.context.pageContext.site.absoluteUrl+"/Lyndon/_api/web/lists/getbytitle('Registration Project')/items";
    const siteURL : string = `${this.context.pageContext.site.absoluteUrl}/Lyndon/_api/web/lists/getbytitle('${encodeURIComponent('Registration Project')}')/items`;


    const itemBody:any = {
      "Title": userNameValue,
      "Email": emailValue,
      "Batch": batchValue,
      "LevelofKnowledge": levelOfKnowledgeValue
    }


    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(itemBody)
    }

    this.context.spHttpClient.post(siteURL, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response:SPHttpClientResponse) =>{
      if(response.status ===201){
        const statusmessage:Element = document.getElementById("divStatus") as HTMLDivElement;
        statusmessage.innerHTML = "Item Created Successfully";
      }else{
        const statusmessage:Element = document.getElementById("divStatus") as HTMLDivElement;
        statusmessage.innerHTML = "An error has occurred " + response.status+" - "+ response.statusText;
      }
    })

  }
  private  _bindAllEvents(): void {
    // create item
    this.domElement.querySelector('#btnCreate')?.addEventListener('click', () =>{
      this.CreateItem();
    })

    // read item
    this.domElement.querySelector('#btnRead')?.addEventListener('click', () =>{
      this.readItems();
    })
  }

  // read items
  private _getListItems(): Promise<IEventRegistration[]>{
    
    const siteURL : string = `${this.context.pageContext.site.absoluteUrl}/Lyndon/_api/web/lists/getbytitle('${encodeURIComponent('Registration Project')}')/items`;
    
    return this.context.spHttpClient.get(siteURL, SPHttpClient.configurations.v1)
    .then((response)=>{
      return response.json();
    }).then((json) =>{
      return json.value
    })as Promise<IEventRegistration[]>;
  }

  // display read items
  private readItems():void{
    this._getListItems().then((listItems) =>{
      let html: string = '<table border=1 width=100% style="border-collapse:collapse">' +
      '<thead>' +
          '<tr>'+
            '<th>ID</th>' +
            '<th>User Name</th>' +
            '<th>Email</th>' +
            '<th>Batch</th>' +
            '<th>Level of Knowledge</th>' +
          '</tr>' +
        '</thead>';
      listItems.forEach((listItem)=>{
        html +=`
        <tr>
        <td>${listItem.Id}</td>
        <td>${listItem.Title}</td>
        <td>${listItem.Email}</td>
        <td>${listItem.Batch}</td>
        <td>${listItem.LevelofKnowledge}</td>
        </tr>
        `;
        html += '</table>';
        const listContainer: Element = document.getElementById("listItems") as HTMLDivElement;
        listContainer.innerHTML = html;
      })

    })
  }


  // functions end


  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.welcome}">
    <table border="4">
        <tr>
            <td><b>Enter User ID</b></td>
            <td><input type="text" id="txtID"></td>
            <td>
                <input type="submit" id="btnSingleItemRead" value="Read Registered User Info">
            </td>
        </tr>
        <tr>
            <td>User Name</td>
            <td><input type="text" id="txtUserName"></td>
        </tr>
        <tr>
            <td>Email</td>
            <td><input type="email" id="txtEmail"></td>
        </tr>
        <tr>
            <td>Enter Batch</td>
            <td>
                <select name="Batch" id="ddlBatch">
                    <option value="Batch 1">Batch 1</option>
                    <option value="Batch 2">Batch 2</option>
                    <option value="Batch 3">Batch 3</option>
                </select>
            </td>
        </tr>
        <tr>
            <td>Enter Level of Knowledge</td>
            <td>
                <select name="LevelOfKnowledge" id="ddlLevelOfKnowledge">
                    <option value="Beginner">Beginner</option>
                    <option value="Intermediate">Intermediate</option>
                    <option value="Expert">Expert</option>
                </select>
            </td>
        </tr>
        <tr>
            <td>
                <input type="submit" value="Create" id="btnCreate">
                
                <input type="submit" value="Read" id="btnRead">
                
                <input type="submit" value="Update" id="btnUpdate">
                
                <input type="submit" value="Delete" id="btnDelete">
            </td>
        </tr>
    </table>
    <div id="divStatus"></div>
  </div>
  <br>
  <hr>
  <div id="listItems">
  </div>
  `;
  this._bindAllEvents();
  }

  // protected onInit(): Promise<void> {
  //   return this._getEnvironmentMessage().then(message => {
  //     this._environmentMessage = message;
  //   });
  // }



  // private _getEnvironmentMessage(): Promise<string> {
  //   if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
  //     return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
  //       .then(context => {
  //         let environmentMessage: string = '';
  //         switch (context.app.host.name) {
  //           case 'Office': // running in Office
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
  //             break;
  //           case 'Outlook': // running in Outlook
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
  //             break;
  //           case 'Teams': // running in Teams
  //           case 'TeamsModern':
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
  //             break;
  //           default:
  //             environmentMessage = strings.UnknownEnvironment;
  //         }

  //         return environmentMessage;
  //       });
  //   }

  //   return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  // }

  // protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
  //   if (!currentTheme) {
  //     return;
  //   }

  //   this._isDarkTheme = !!currentTheme.isInverted;
  //   const {
  //     semanticColors
  //   } = currentTheme;

  //   if (semanticColors) {
  //     this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
  //     this.domElement.style.setProperty('--link', semanticColors.link || null);
  //     this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
  //   }

  // }

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
