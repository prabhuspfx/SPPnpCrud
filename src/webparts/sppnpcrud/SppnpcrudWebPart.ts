import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SppnpcrudWebPart.module.scss';
import * as strings from 'SppnpcrudWebPartStrings';

import { sp } from '@pnp/sp'

export interface ISppnpcrudWebPartProps {
  description: string;
}

export interface ISPList {    
  ID: string;    
  ProfileName: string;    
  ProfileJob: string;    
}   

export default class SppnpcrudWebPart extends BaseClientSideWebPart<ISppnpcrudWebPartProps> {



  private AddEventListeners() : void{    
   
    document.getElementById('AddSPItem').addEventListener('click',()=>this.AddSPItem());    
    document.getElementById('UpdateSPItem').addEventListener('click',()=>this.UpdateSPItem());    
    document.getElementById('DeleteSPItem').addEventListener('click',()=>this.DeleteSPItem());    
   }    
       
    private _getSPItems(): Promise<ISPList[]> {    
    return sp.web.lists.getByTitle("ProfileList").items.get().then((response) => {    
          
       return response;    
     });    
            
   }    
       
    private getSPItems(): void {    
          
       this._getSPItems()    
         .then((response) => {    
           this._renderList(response);    
         });    
   }    
       
   private _renderList(items: ISPList[]): void {    
     let html: string = '<table class="TFtable" border=1 width=style="bordercollapse: collapse;">';    
     html += `<th></th><th>ProfileId</th><th>Name</th><th>Job</th>`;    
     if (items.length>0)  
     {  
     items.forEach((item: ISPList) => {    
       html += `    
            <tr>   
            <td>  <input type="radio" id="ProfileId" name="ProfileId" value="${item.ID}"> <br> </td>   
            
           <td>${item.ID}</td>    
           <td>${item.ProfileName}</td>    
           <td>${item.ProfileJob}</td>    
           </tr>    
           `;     
     });    
    }  
    else    
    
    {  
      html +="No records...";  
    }  
     html += `</table>`;    
     const listContainer: Element = this.domElement.querySelector('#DivGetItems');    
     listContainer.innerHTML = html;    
   } 

  public render(): void {
    this.domElement.innerHTML = `    
     <div class="parentContainer" style="background-color: white">    
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">    
       <div class="ms-Grid-col ms-u-lg   
   ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">   
     
           
       </div>    
    </div>    
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">    
       <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:   
   x;">Profile Details</div>    
           
    </div>    
    <div style="background-color: white" >    
       <form >    
          <br>    
          <div data-role="header">    
             <h3>Add SharePoint List Items</h3>    
          </div>    
           <div data-role="main" class="ui-content">    
             <div >    
                
               
               <input id="ProfileName"  placeholder="ProfileName"/>    
               <input id="ProfileJob"  placeholder="ProfileJob"/>    
               <button id="AddSPItem"  type="submit" >Add</button>    
               <button id="UpdateSPItem" type="submit" >Update</button>    
               <button id="DeleteSPItem"  type="submit" >Delete</button>  
             </div>    
           </div>    
       </form>    
    </div>    
    <br>    
    <div style="background-color: white" id="DivGetItems" />    
      
    </div>    
       
    `;    
 this.getSPItems();    
 this.AddEventListeners(); 
  }

  protected AddSPItem()    
 {      
     
      sp.web.lists.getByTitle('ProfileList').items.add({        
        ProfileName : document.getElementById('ProfileName')["value"],    
        ProfileJob : document.getElementById('ProfileJob')["value"]  
         
     });   
   
      alert("Record with Profile Name : "+ document.getElementById('ProfileName')["value"] + " Added !");    
        
 }  
 
 protected UpdateSPItem()    
 {      
  var ProfileId =  this.domElement.querySelector('input[name = "ProfileId"]:checked')["value"];  
     sp.web.lists.getByTitle("ProfileList").items.getById(ProfileId).update({    
      ProfileName : document.getElementById('ProfileName')["value"],    
      ProfileJob : document.getElementById('ProfileJob')["value"]  
        
   });    
  alert("Record with Profile ID : "+ ProfileId + " Updated !");    
 }  

 protected DeleteSPItem()    
 {      
  var ProfileId =  this.domElement.querySelector('input[name = "ProfileId"]:checked')["value"];  
    
      sp.web.lists.getByTitle("ProfileList").items.getById(ProfileId).delete();    
      alert("Record with Profile ID : "+ ProfileId + " Deleted !");    
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
