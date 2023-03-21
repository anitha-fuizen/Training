import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'News1WebPartStrings';
import styles from './components/News1.module.scss';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import "@pnp/sp/sputilities";
//import { IHttpClientOptions, HttpClientResponse, HttpClient } from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { getSP } from './components/pnpConfig';
import { SPFI} from '@pnp/sp';
export interface IGetListItemFromSharePointListWebPartProps {

  description: string;
}
export interface ISPLists
{
  value: ISPList[];
}
export interface ISPList
{
  Title: string;
  Description:String;
  
}
export interface Igetdetails{
  Reportingmanager:{
    Email:string;
    
  }
}
export default class GetListItemFromSharePointListWebPart extends BaseClientSideWebPart <IGetListItemFromSharePointListWebPartProps> {
constructor(props:any)
{
  super();
 // this.send_Email=this.send_Email.bind(this);
 
}
  private _getListData(): Promise<ISPLists>
  {
   return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('Training')/Items?$select=Title,Description",
       SPHttpClient.configurations.v1

   )
   .then((response: SPHttpClientResponse) =>
       {
       return response.json();
        console.log(response.json())
       });
   }
   private _renderListAsync(): void
   {
    if (Environment.type === EnvironmentType.SharePoint ||
             Environment.type === EnvironmentType.ClassicSharePoint) {
     this._getListData()
       .then((response) => {
         this._renderList(response.value);
         console.log(response.value);
       }).catch((err)=>{console.log(err)})
}
 }
 private  _getData(useraddress:any): Promise<Igetdetails> {
  return   this.context.spHttpClient
     .get(
       this.context.pageContext.web.absoluteUrl +
       "/_api/web/lists/GetByTitle('Employee List')/Items?$select=Reportingmanager/EMail&$expand=Reportingmanager&$filter = employeename/EMail eq "+useraddress,
       SPHttpClient.configurations.v1

     ) .then(async (response: SPHttpClientResponse) => {
      //console.log(response)
 
      const resItems= response.json();
   
      
    return resItems;
     //then
    
      
    })
   
}





 
  



 public async send_Email(title:any):Promise<void>{
    
  try{
  let _sp:SPFI = getSP(this.context);
 
  let addressString: string = await _sp.utility.getCurrentUserEmailAddresses();
  //let mymanager:string= await _sp.web.lists.getByTitle("EmployeeDetails").items.getItemByStringId() 
//const items:any=   await _sp.web.lists.getByTitle("Employee List").items.select("Reportingmanager/Title").expand("Reportingmanager")();
console.log(addressString)
let emailString:any=null;
const mymanager=this._getData(addressString) 
mymanager.then((x:any)=>{
    
  let obj = x.value;
 
 obj.forEach((x:any)=>{
    emailString = x.Reportingmanager.EMail;
    console.log(emailString);
     _sp.utility.sendEmail({
      To: [emailString],
     
      Subject: "Request for"+title,
      Body: "Iam interested in "+title,
      AdditionalHeaders: {
          "content-type": "text/html"
      },
  });
 })
 return emailString
})

 
//return items
window.alert("success")
console.log("emailsend");
  }
  catch(e){
     console.log(e);
  }
 }
 private async _renderList(items: ISPList[]): Promise<void>
 {

  
let  html: string = '<table border=2 width=80% style="font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;>';
 let x=0;
  html+=`<th>Title</th><th>Description</th><th>Apply</th>`
   console.log(items)
  items.forEach((item: ISPList) => {
x=x+1;
    
    html += `<tr>

       <td><label> ${item.Title}</label></td>
       <td> ${item.Description}</td>

       <td> <button id="nominate_btn${x}" class="nominate1">Nominate</button></td><br><br>
       
        
        </tr> `;
       
       
  });
html += "</table>"; 


  const listContainer: Element = this.domElement.querySelector('#BindspListItems');
  listContainer.innerHTML = html;
  
  // let clickEvent: Element =  this.domElement.querySelector('.nominate1');
   //clickEvent.addEventListener("click", (e: Event) => this.sendEmail());
   //clickEvent=null;

  var buttons = listContainer.getElementsByTagName("BUTTON");
  var labels=listContainer.getElementsByTagName("LABEL");
 if(buttons)
  {
 for(let i=0;i<buttons.length;i++) {   buttons[i].addEventListener('click', (e: Event) => this.send_Email(labels[i].textContent)); }
  }
}


  public render(): void {
    this.domElement.innerHTML = `
      <div class={styles.sharepointframe}>
    <div class={ styles.container }>
      <div class={ styles.row }>
        <div class={ column }>
        <span class="${styles.title}"></span>
          
          </div>
          <br/>
          <br/>
          <br/>
          <div id="BindspListItems" />
          </div>
          </div>
           
          </div>`;
          this._renderListAsync();
          this._getData("sindhu");
  
          
        
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