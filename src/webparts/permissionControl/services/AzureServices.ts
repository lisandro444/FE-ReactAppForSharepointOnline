import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { IHttpClientOptions, HttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";
import { bool } from "prop-types";

interface IPermissionResponse{
  UserHasEnumeratePermission : boolean;
  DisplayAddRemoveButton : boolean;
  IsWebUnique : boolean;
  UsersToBeAdded : string;
  GroupId : Number;
  InvalidUsers: string;
  IsSuccess : boolean;
  Message : string;      
}

export class AzureService {
    //context: any;
    constructor() {

    }
    
   
 public async sharingCapabilityAzureFunction(context:BaseWebPartContext)  {  
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
       
    let returnObject;
    const postOptions: IHttpClientOptions = {
        headers: requestHeaders,
        body: `{ SiteUrl: '${context.pageContext.web.absoluteUrl}', LoginName: '${context.pageContext.user.loginName}'} `
      };
      
    const promise = context.httpClient
    .post('https://scj-role-def.azurewebsites.net/api/scj-role-def', HttpClient.configurations.v1, postOptions)
    .then((response: HttpClientResponse) => {
      returnObject = response.json();
      })
    .catch(err=>{
      console.log(err);
    })

    await promise;
    return returnObject;
    //return JSON.stringify(returnObject);
    
 }

 public async DoesWebHaveUniquePermissions(webUrl,context:BaseWebPartContext){
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    
    let returnObject;
    const postOptions: IHttpClientOptions = {
        headers: requestHeaders,
        body: `{ SiteUrl: '${webUrl}', Module: 'CheckWebInheritance'} `
      };
      
    const promise = context.httpClient
      .post('https://scj-role-def.azurewebsites.net/api/scj-role-def', HttpClient.configurations.v1, postOptions)
      .then((response: HttpClientResponse) => {
        returnObject = response.json();
        })
      .catch(err=>{
        console.log(err);
      })

    await promise;
    return returnObject;
    
  }
  public async AddGroup(webUrl, emails, groupName,context:BaseWebPartContext){
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    
    let returnObject;
    const postOptions: IHttpClientOptions = {
        headers: requestHeaders,
        body: `{ SiteUrl: '${webUrl}', Module: 'UserGroup', Emails:'${emails}', GroupName:'${groupName}'} `
      };
      
    const promise = context.httpClient
      .post('https://permissioncontrolapi20190825023420.azurewebsites.net/api/permission-control', HttpClient.configurations.v1, postOptions)
      .then((response: HttpClientResponse) => {
        returnObject = response.json();       
        })
      .catch(err=>{
        console.log(err);
      })

    await promise;
    return returnObject;
    
  }


  public AddGroup1(webUrl, emails, groupName,context:BaseWebPartContext) :Promise<any>{
    return new Promise<IPermissionResponse>((resolve:(permissionControlResponse:IPermissionResponse) => void, reject:(error:any)=>void):void=>{

      const requestHeaders: Headers = new Headers();
      requestHeaders.append("Content-type", "application/json");
      requestHeaders.append("Cache-Control", "no-cache");

      const postOptions: IHttpClientOptions = {
        headers: requestHeaders,
        body: `{ SiteUrl: '${webUrl}', Module: 'UserGroup', Emails:'${emails}', GroupName:'${groupName}'} `
      };

      context.httpClient
      .post('https://permissioncontrolapi20190825023420.azurewebsites.net/api/permission-control', HttpClient.configurations.v1, postOptions)
      .then((response: HttpClientResponse) : Promise<IPermissionResponse> =>{
        return response.json();
      })
      .then((permissionControlResponse:IPermissionResponse) : void =>{
        resolve(permissionControlResponse);
      })
      ,(error:any)=>{
        reject(error);
      }
    });
  }

}

