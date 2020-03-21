import { INavLinkGroup, INavLink } from 'office-ui-fabric-react/lib/Nav';
import { Web, sp, Folder, Folders } from '@pnp/sp';
import {IDetailsListItem} from  '../components/IGridControlProps'
import {ICustomNavLink, LinkType} from '../components/INavControlProps'
import { AzureService } from './AzureServices';
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from '@pnp/sp';

export class PnPService {
    private _azureService;
    private _context;
    constructor(context:BaseWebPartContext) {
        this._azureService = new AzureService();
        this._context = context;
    }
   
    public async getWeb(url): Promise<INavLinkGroup[]> {
        let groups: INavLinkGroup[] = [];
        try{
            let initialweb = new Web(url);
            let web = await initialweb.get();
            let siteList = new Array<ICustomNavLink>();
            let subLinks = await this.getSubSitesAndLists(url);
            siteList.push({
                name: web.Title,
                url: '#',
                target: '_blank',
                isExpanded: false,
                links: subLinks,
                icon: "",
                customUrl:url,
                webUrl:url,
                isUniquePermissions:false,
                linkType:LinkType.web   
            });

            
            groups.push(
                {
                    links: siteList
                }
            );       
            return groups;
        }
        catch(error){
            console.log("Get WEb: "+ error);
            return groups;
        }
    }
    // Get subwebs and lists with broken inheritance
    private async getSubSitesAndLists(url): Promise<INavLink[]> {
        let childrenWebs = new Array<ICustomNavLink>();
        try{         

            let web = new Web(url);
            try{
                let listswithBrokenInheritance = await this.getListsWithBrokenInheritance(url);
                listswithBrokenInheritance.forEach( list => {
                    this.getUniqueFolders(list.RootFolder, web, list.webUrl).then(childFolders=>{
                        if(list.HasUniqueRoleAssignments ||
                            childFolders.length > 0){
                                childrenWebs.push(
                                    {
                                        name: list.Title,
                                        url: '#',
                                        target: '_blank',
                                        isExpanded: false,
                                        links: childFolders,
                                        icon: list.HasUniqueRoleAssignments?'BranchLocked':'',
                                        customUrl:url,
                                        webUrl:list.webUrl,
                                        isUniquePermissions: true,
                                        linkType:LinkType.list   
                                    }
                                );
                            }
                    });
                
                });
            }
            catch(listError){
                console.log("getSubSitesAndLists-list error"+listError)
            }

            try{
               //let webs = await web.getSubwebsFilteredForCurrentUser().expand("HasUniqueRoleAssignments").get();
                let webs = await web.getSubwebsFilteredForCurrentUser().get();

                await webs.forEach(async r => {
                    try{
                        let webUrl =  window.location.protocol + "//" + window.location.host+r.ServerRelativeUrl; 
                        let childLinks = await this.getSubSitesAndLists(webUrl);
                        let response = 
                               await this._azureService.DoesWebHaveUniquePermissions(webUrl,this._context);

                        var isWebUnique  = JSON.parse(response).IsWebUnique;
                       
                        let childWeb = {
                            name: r.Title,
                            url: '#',
                            target: '_blank',   
                            isExpanded: false,
                            links: childLinks,
                            //icon: r.HasUniqueRoleAssignments?'BranchLocked':'',
                            icon:isWebUnique ? 'BranchLocked':'',
                            customUrl: webUrl,
                            webUrl: webUrl,
                           // isUniquePermissions: r.HasUniqueRoleAssignments ? true:false,
                            isUniquePermissions: isWebUnique ? true:false,
                            linkType:LinkType.web   
                                        
                        };
                        childrenWebs.push(childWeb);
                    }
                    catch(e){
                        console.log("getSubSitesAndLists-recursive web error"+e);
                    }
                });
             }
             catch(ex){
                console.log("getSubSitesAndLists-web error"+ex) ;
             }

            return await childrenWebs;
        }
        catch(error){
            console.log("getSubSitesAndLists : " + error );
            return childrenWebs;
        }
    }

    

    private async getListsWithBrokenInheritance(webUrl: string) {
        let brokenInheritanceListArray = [];
        try{
            let brokenInheritanceListArray = [];
            let currentweb = new Web(webUrl);
            let lists = await currentweb.lists
            .filter(`Hidden eq false and Title ne 'Site Pages' and Title ne 'Site Assets' and Title ne 'Style Library' and Title ne 'Form Templates'`)
            .expand("RootFolder")
            .select("HasUniqueRoleAssignments, Title, RootFolder")
            .get();

            lists.forEach(listElement => {

                brokenInheritanceListArray.push({
                    Title: listElement.Title,
                    RootFolder:listElement.RootFolder,
                    HasUniqueRoleAssignments:listElement.HasUniqueRoleAssignments,
                    webUrl: webUrl    
                })
            });
            return brokenInheritanceListArray;
        }
        catch(error){
            console.log("getListsWithBrokenInheritance " + error);
            return brokenInheritanceListArray;
        }
    }

    private async getUniqueFolders(parentFolder, web:Web, webUrl:String):Promise<Array<INavLink>>{
        var childFolders: INavLink[] = [];       
        try{
            let folders = await web
            .getFolderByServerRelativeUrl(parentFolder.ServerRelativeUrl)
            .folders
            .filter(`Name ne 'Attachments' and Name ne 'Item' and Name ne 'Forms' and Name ne 'InfoPath Form Template'`)      
            .expand("ListItemAllFields/HasUniqueRoleAssignments")     
            .select('Name,ServerRelativeUrl,ListItemAllFields/HasUniqueRoleAssignments')       
            .get();       
        
            const promises = folders.map(async(folder)=>{
                
                let childLinks = await this.getUniqueFolders(folder, web, webUrl);
            //  console.table(folder);
                if((folder.ListItemAllFields && folder.ListItemAllFields.HasUniqueRoleAssignments == true)
                        || childLinks.length != 0){
                        let childFolder = {
                            name: folder.Name,
                            url: '#',
                            target: '_blank',
                            isExpanded: false,
                            links: childLinks,
                            icon:  folder.ListItemAllFields.HasUniqueRoleAssignments?'BranchLocked':'',
                            customUrl:folder.ServerRelativeUrl,
                            webUrl:webUrl,
                            isUniquePermissions: folder.ListItemAllFields.HasUniqueRoleAssignments?true:false,
                            linkType:LinkType.folder   
                        };

                childFolders.push(childFolder);
                }
            });      
            await Promise.all(promises);
            return childFolders; 
         }    
         catch(error){
            console.log("getUniqueFolders-"+error);
            return childFolders; 
         }
    
    }

    public async getUsersForWeb(customNavLink:ICustomNavLink): Promise<Array<IDetailsListItem>> {        
        let rolesColl = new Array<IDetailsListItem>();
        try{                    
            var counter=0;

            const web = new Web(customNavLink.webUrl);

            try{
                var roleAssignments = await this.getRoleAssignmentsForNavItem(customNavLink);
                rolesColl = await this.getRolesCollection(roleAssignments,customNavLink);
            }
            catch(error){
               console.log(error)
               rolesColl = await this.getRolesCollectionForNonAdmin(customNavLink);
            }

            return rolesColl;
        }
        catch(exceptions)
        {
            console.log("get users " + exceptions);
        }
    }  

    private async getRoleAssignmentsForNavItem(customNavLink:ICustomNavLink){
        const web = new Web(customNavLink.webUrl)
        
        let roleAssignments;  

        try{
        
            switch(customNavLink.linkType) { 
                case LinkType.web: {               
                    roleAssignments = await web.roleAssignments
                    .expand("Member", "RoleDefinitionBindings").get();      
                break; 

                } 
                case LinkType.list: { 
                    roleAssignments = await web.lists
                    .getByTitle(customNavLink.name)
                    .roleAssignments.expand("Member", "RoleDefinitionBindings").get();
                break; 

                } 
                case LinkType.folder: { 
                    roleAssignments = await web.getFolderByServerRelativeUrl(customNavLink.customUrl)
                    .expand("ListItemAllFields/RoleAssignments/Member","ListItemAllFields/RoleAssignments/RoleDefinitionBindings")
                    .get();

                    roleAssignments = roleAssignments.ListItemAllFields.RoleAssignments;
                    break; 

                }
                default: { 
                //statements; 
                break; 

                } 
            } 

            await roleAssignments;
            return roleAssignments;
        }
        catch(error){
            console.log("getRoleAssignmentsForNavItem-" + error);
            throw(error);
        }
    }

    private async getRolesCollection(roleAssignments,customNavLink){

        let rolesColl = new Array<IDetailsListItem>(); 

        var counter=0;

        const web = new Web(customNavLink.webUrl);

        const promises = roleAssignments.map( async(roleAssignment) => { 

            if(roleAssignment.Member["odata.type"] == "SP.Group") {
                
                if(roleAssignment.Member["Title"].toLowerCase().includes("external"))
                {
                    /*var grps = await this.getExternalUserGroups(customNavLink.webUrl);
                    grps.forEach( grp =>{
                        let objUser = { 
                            key: counter,              
                            name: String(grp.Title),
                            permission: '',
                            spgroup: '',
                            userandgroup: String(grp.LoginName)
                        };
                        rolesColl.push(objUser);
                        counter++;
                    })*/
                    let objUser = { 
                        key: counter,              
                        name: '',
                        permission: String(roleAssignment.RoleDefinitionBindings[0].Name),
                        spgroup: String(roleAssignment.Member.LoginName),
                        userandgroup:''
                    };
                    rolesColl.push(objUser);
                    counter++;
                } 
                else{
                    let grp = await web.siteGroups.getById(roleAssignment.Member.Id)
                        .expand("Users")
                        .select("Id", "LoginName", "PrincipalType","Title", "IsSiteAdmin")
                        .get();
                        
                    grp.Users.forEach((u:any) => {             
                        let objUser = { 
                                key: counter,              
                                name: String(u.Title),
                                permission: String(roleAssignment.RoleDefinitionBindings[0].Name),
                                spgroup: String(roleAssignment.Member.LoginName),
                                userandgroup: String(u.LoginName)
                            };
                        // Append collection
                    if (this.goodUser(u.LoginName) && u["PrincipalType"] != "1"
                            && u["Title"].toLowerCase() != "company administrator"
                            && u["Title"].toLowerCase() != "everyone except external users") {
                            rolesColl.push(objUser);
                            //return rolesColl;
                            }
                        counter++;
                        });
                    }
            } 

        });      
        await Promise.all(promises);
        return rolesColl;
    }

    private async getRolesCollectionForNonAdmin(customNavLink):Promise<Array<IDetailsListItem>> {
        let rolesColl = new Array<IDetailsListItem>(); 
        const oWeb = new Web(customNavLink.webUrl);
        var counter =0;
        oWeb.siteGroups      
        .get()
        .then(groups => {
          // Flatten array[], first RoleDef{}
          
          groups.forEach(group => {   
                     
            oWeb.siteGroups
              .getById(group.Id)
              .users.get()
              .then(users => {
                // Loop Group members
                users.forEach(u => {                     
                  let obj = {
                    key:counter,
                    name: u.Title,                
                    permission: "Please contact site owner",//row.RoleDef,
                    //EARS: u.Title,                   
                    spgroup: group.LoginName,
                    userandgroup: u.LoginName
                  };  
                    // Append collection
                    if (this.goodUser(u.LoginName) && u["PrincipalType"] != "1"
                         && u["Title"].toLowerCase() != "company administrator"
                             && u["Title"].toLowerCase() != "everyone except external users") {
                    rolesColl.push(obj);
                    counter++;
                        //}
                    }
                }); 
                
              })
              .catch((e)=>{
                console.log(e);     
              });  
          });                 
        })        
        .catch(e =>
        {
          console.log(e);       
        }); 

        return await rolesColl;  
    }

   

    private async getExternalUserGroups(webUrl:string){
        const web = new Web(webUrl);
        var groups = await web.siteGroups
        .filter(`Title eq 'SCJ External Contributor' or Title eq 'SCJ External Reader' `)
        .get();      

        return await groups;
    }

    private goodUser(userLoginName:string) {
    // Colon Display Name
    if (userLoginName && userLoginName.indexOf('|') != -1) {   
        let arrayUserAndGroupDetails = userLoginName.split('|');
        if(arrayUserAndGroupDetails[1].toLowerCase() === "tenant"){
            return true;        
        }
    }
    return false;
    }

    public async getPermissionsForList(navLink:ICustomNavLink){
        try{
            const web = new Web(navLink.customUrl);
            var listRA = await web.lists.getByTitle(navLink.name).roleAssignments.expand("Member", "RoleDefinitionBindings").get();
        }
        catch(error)
        {
            console.log("Get PermissionFor List"+error);
        }      
    }

    public async addExternalUsers(externalUserEmails:string, groupID:string, siteUrl:string) {
            
        const emailBody = 'Welcome to site';
        var externalEmailsArray = externalUserEmails.split(',');
        const promises = externalEmailsArray.map(externalUserEmail => {
            const requestHeaders: Headers = new Headers();
            requestHeaders.append("Content-type", "application/json");
            requestHeaders.append("Cache-Control", "no-cache");
            const client = new SPHttpClient();
            client.post(`${siteUrl}/_api/SP.Web.ShareObject`, {
              body: JSON.stringify({
                emailBody,
                includeAnonymousLinkInEmail: false,
                peoplePickerInput: JSON.stringify([{
                  Key: externalUserEmail,
                  DisplayText: externalUserEmail,
                  IsResolved: true,
                  Description: externalUserEmail,
                  EntityType: '',
                  EntityData: {
                    SPUserID: externalUserEmail,
                    Email: externalUserEmail,
                    IsBlocked: 'False',
                    PrincipalType: 'UNVALIDATED_EMAIL_ADDRESS',
                    AccountName: externalUserEmail,
                    SIPAddress: externalUserEmail,
                    IsBlockedOnODB: 'False'
                  },
                  MultipleMatches: [],
                  ProviderName: '',
                  ProviderDisplayName: ''
                }]),
                roleValue: "group:" + groupID , // where `6` is a GroupId
                sendEmail: true,
                url: siteUrl,
                useSimplifiedRoles: true
              })
            })
              .then(r => { console.log(r.json); })
              .catch(e => { console.log(e); });
        });
        await Promise.all(promises);

      }
}