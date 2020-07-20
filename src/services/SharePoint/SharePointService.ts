import { WebPartContext } from "@microsoft/sp-webpart-base";
import { EnvironmentType } from "@microsoft/sp-core-library";
import { SPHttpClient } from "@microsoft/sp-http";
import { IListCollection } from "./IList";
import { IListFieldCollection } from "./IListField";
import { IListItemCollection } from "./IListItem";

export class SharePointServiceManager {
    public context: WebPartContext;
    public environmentType: EnvironmentType;
    public itemID: number;
    public ideaListID: string;

    public setup(context: WebPartContext, environmentType: EnvironmentType): void {
        this.context = context;
        this.environmentType = environmentType;
    }

    public get(relativeEndpointUrl: string): Promise<any> {
        //console.log(`${this.context.pageContext.web.absoluteUrl}${relativeEndpointUrl}`);
        return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}${relativeEndpointUrl}`, SPHttpClient.configurations.v1)
        .then(
            response => {
                return response.json()
            }
        )
        .catch(error => {
            return Promise.reject(error);
        });
    }

    public getLists(): Promise<IListCollection> {
        return this.get('/_api/lists');
    }

    public getListItems(listId: string, selectedFields?: string[]) : Promise<IListItemCollection>{
        return this.get(`/_api/lists/getbyid('${listId}')/items?$select=*,Author/Name,Author/Title,LinkToSpec/Title&$expand=Author/Id,LinkToSpec/Id,AttachmentFiles`);
    }

    public getListItem(listId: string, itemId: number){
        return this.get(`/_api/lists/getbyid('${listId}')/items(${itemId})?$select=*,Author/Name,Author/Title,Author/EMail,LinkToSpec/Title&$expand=Author/Id,LinkToSpec/Id,AttachmentFiles`);
    }

    public getListItemVersions(listId: string, itemId: number){
        return this.get(`/_api/lists/getbyid('${listId}')/items(${itemId})/versions?$select=*,Author/Name,Author/Title,LinkToSpec/Title&$expand=Author/Id,LinkToSpec/Id,AttachmentFiles&$orderby=Created asc`);
    }

    public getListItemsFIltered(listId: string, filterString: string) : Promise<IListItemCollection>{
        //console.log(`/_api/lists/getbyid('${listId}')/items?$filter=IdeaStatus eq '${filterString}'`);
        return this.get(`/_api/lists/getbyid('${listId}')/items?$select=*,Author/Name,Author/Title,LinkToSpec/Title&$expand=Author/Id,LinkToSpec/Id,AttachmentFiles&$filter=IdeaStatus eq '${filterString}'`);
    }
    

    public getListFields(listId: string, showHiddenField: boolean = false): Promise<IListFieldCollection>{
        return this.get(`/_api/lists/getbyid('${listId}')/fields${!showHiddenField ? '?$filter=Hidden eq false' : ''}`);
    }

    
    public getUserByID(userID: string): Promise<any> {
        return this.get(`/_api/web/getuserbyid(${userID})`);
    }

    
    
    
    public getUsers(): Promise<any> {
        return this.get(`/_api/web/siteusers`);
    }

    public getSoftwareDeveloperGroupMembers() {
        return this.get(`/_api/Web/SiteGroups/GetByName('SoftwareDeveloper')/users`);
    }

    public getUserGroupMembers() {
        return this.get(`/_api/Web/SiteGroups/GetByName('User')/users`);
    }

    public getAdminGroupMembers() {
        return this.get(`/_api/Web/SiteGroups/GetByName('Admin')/users`);
    }

    public getGroupsOfCurrentUser(): Promise<any> {
        return this.get(`/_api/web/currentuser/groups`);
    }

    public changeStatus(relativeEndpointUrl: string, newStatus: string) {
        
        return this.context.spHttpClient.fetch(`${this.context.pageContext.web.absoluteUrl}${relativeEndpointUrl}`, SPHttpClient.configurations.v1,{
            
            method: "PATCH",
            
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': '',
                'if-match': '*',
            },
            
            body : JSON.stringify({
                '__metadata': {
                    'type': 'SP.Data.IdeaListItem'
                },
                "IdeaStatus": newStatus

            })
        })
        .then(
            response => {
                return response.status;
            }
        )
        .catch(error => {
            return Promise.reject(error);
        });
    }

    public addUserToSoftwareDeveloperGroup(loginName) {
        const body = JSON.stringify({
            '__metadata': {
                'type': 'SP.User'
            },
            'LoginName': loginName
        })

        return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/Web/SiteGroups/GetByName('SoftwareDeveloper')/users`, SPHttpClient.configurations.v1,
        {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': ''
            },
            body: body
        })
        .then(
            response => {
                return response.json()
            }
        )
        .catch(error => {
            return Promise.reject(error);
        });
    }


    public updateIdea(name, desc, formula, status){

        const body = JSON.stringify({
            '__metadata': {
                'type': 'SP.Data.IdeaListItem'
            },
            'Title': name,
            'Comment1': desc,
            'IdeaFormula': formula,
            "IdeaStatus": status
        })
        //console.log(name);
        //console.log(desc);
        //console.log(formula);
        //console.log(this.context.pageContext.web.absoluteUrl);
        return this.context.spHttpClient.fetch(`${this.context.pageContext.web.absoluteUrl}/_api/lists/getbyid('${this.ideaListID}')/items(${this.itemID})`, SPHttpClient.configurations.v1,
        {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': '',
                'if-match': '*',
            },
            method: "PATCH",
            body: body
        })
        .then(
            response => {
                return response.status;
            }
        )
        .catch(error => {
            return Promise.reject(error);
        });
    }
     
    

}

const SharePointService = new SharePointServiceManager();

export default SharePointService;  //singleton pattern