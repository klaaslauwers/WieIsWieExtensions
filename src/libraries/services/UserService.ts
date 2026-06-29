import { IUserService } from "./IUserService";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientFactory, MSGraphClientV3, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { ISearchResult } from "../data/ISearchResult";
import { IUserPresence } from "../data/IUserPresence";
import { PnPClientStorage } from "@pnp/common";
import { IProfileImage } from "../data/IProfileImage";

export interface IGraphBatchResponseBody {
  responses: IGraphPhotoBatchResponse[];
}

export interface IGraphPhotoBatchResponse {
  id: string;
  status: number;
  body: string;
  headers: {
    'Content-Type': string;
  };
}

export interface IGraphBatchRequestBody {
  requests: IGraphBatchRequest[];
}

export interface IGraphBatchRequest {
  id: string;
  method: string;
  url: string;
}

//https://www.eliostruyf.com/fix-admin-consent-sp-token-retrieval-flows-spfx/
export class UserService implements IUserService {
  public static readonly serviceKey: ServiceKey<UserService> =
    ServiceKey.create<UserService>("SPFx:UserService", UserService);

  private _spHttpClient: SPHttpClient;
  private _msGraphClient: MSGraphClientFactory;
  private _pageContext: PageContext;
  private _currentWebUrl: string;
  private storage = new PnPClientStorage();

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
      this._msGraphClient = serviceScope.consume(MSGraphClientFactory.serviceKey);
      this._currentWebUrl = this._pageContext.web.absoluteUrl;
    });
  }

  public async getUserData(searchResultItem: ISearchResult, expandField: string): Promise<ISearchResult> {
    const newResultItem: ISearchResult = {...searchResultItem};

    try {
      const usersresponse: SPHttpClientResponse = await this._spHttpClient.get(
        `${this._currentWebUrl}/_api/web/lists('${newResultItem.IdentityListId}')/items(${newResultItem.ListItemID})?$select=${expandField}/EMail,${expandField}/JobTitle&$expand=${expandField}`,
        SPHttpClient.configurations.v1
      );
      if (usersresponse.ok) {
        const profileData = await usersresponse.json();

        newResultItem.Mail = profileData[expandField]?.EMail ?? '';
        newResultItem.JobTitle = profileData[expandField]?.JobTitle ?? '';

        if(newResultItem.Mail) {
          
          const propsResponse: SPHttpClientResponse = await this._spHttpClient.get(
            `${this._currentWebUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v=%27i%3A0%23.f%7Cmembership%7C${newResultItem.Mail}%27`,
            SPHttpClient.configurations.v1
          )
          if (propsResponse.ok) {
            const jsonUser = await propsResponse.json();
            newResultItem.UserId = this.getUserId(jsonUser.UserProfileProperties); 
          }
        }

        newResultItem.PictureThumbnailURL = `${newResultItem.SPSiteURL}/_layouts/15/userphoto.aspx?size=L&accountname=${newResultItem.Mail}`;
      }
    } catch (error) {
      console.error(error);
    }
    return newResultItem;
  }

  public async getPresenceStatusCurrent(userId: string): Promise<IUserPresence> {
    // Get update status for current user
    try{
      const _currentUserPresenceUpd: IUserPresence[] = await this.getUserPresence([userId]);
      return { activity: _currentUserPresenceUpd[0].activity, availability: _currentUserPresenceUpd[0].availability };
    }
    catch(e){
      console.log(e)
      return { activity: "", availability: ""};
    }
  }

  public async getUserPresence(userObjIds: string[]): Promise<IUserPresence[]> {
    // Get presences for Users Ids
    // TODO - CHECK https://learn.microsoft.com/en-us/graph/sdks/create-client?tabs=typescript
    try{
      return await this._msGraphClient.getClient("3").then(async (client: MSGraphClientV3) => {
        const _presence: any = await client.api(`/communications/getPresencesByUserId`)
          .version("beta")
          .post({ ids: userObjIds });
        console.log("USER PRESENCE: ");
        console.log(_presence);
        return _presence.value;
      });
    }
    catch(e){
      console.log(e);
      return [{ activity: "", availability: ""}];
    }
  }

  public getUserId(userProfileProperties: any[]): string{
    // Get User Properties
    // use type assertion
    const props = {} as { [key: string]: any };
    userProfileProperties.forEach((prop) => {
      props[prop.Key] = prop.Value;
    });
    // Get UserID
    return props["msOnline-ObjectId"];
  }

  public async getImageBase64(users: string[]): Promise<IProfileImage> {
    const graphClient = await this._msGraphClient.getClient('3');
        
    const body: IGraphBatchRequestBody = { requests: [] };
        
    users.forEach((user) => {
      const requestUrl: string = `/users/${user}/photo/$value`;
      body.requests.push({ id: user, method: 'GET', url: requestUrl });
    });

    const response: IGraphBatchResponseBody = await graphClient.api('$batch').version('v1.0').post(body);

    const results: IProfileImage = {};
    response.responses.forEach(r => {
      if (r.status === 200) {
        results[r.id] = `data:${r.headers["Content-Type"]};base64,${r.body}`;
      }
    });
    
    return results;
  }

  public async getImageBase64FromUrl(pictureUrl: string, userId: string): Promise<string> {
    return new Promise((resolve, reject) => {
      const value = this.storage.local.get(userId);
      if (value) {
        resolve(value);
      }
      const image = new Image();
      image.addEventListener("load", () => {
        const tempCanvas = document.createElement("canvas");
        (tempCanvas.width = image.width),// eslint-disable-line
          (tempCanvas.height = image.height),
          tempCanvas.getContext("2d")!.drawImage(image, 0, 0);
        let base64Str;
        try {
          base64Str = tempCanvas.toDataURL("image/png");
        } catch (e) {
          return "";
        }
        this.storage.local.put(userId, base64Str);
        resolve(base64Str);
      });
      image.src = pictureUrl;
    });
  }
}