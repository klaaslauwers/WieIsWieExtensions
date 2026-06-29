import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { ISearchResult } from "../data/ISearchResult";
import { ISharePointService } from "./ISharePointService";

export class SharePointService implements ISharePointService {
  public static readonly serviceKey: ServiceKey<SharePointService> =
    ServiceKey.create<SharePointService>("SPFx:SharePointService", SharePointService);

  private _spHttpClient: SPHttpClient;
  private _pageContext: PageContext;
  private _currentWebUrl: string;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
      this._currentWebUrl = this._pageContext.web.absoluteUrl;
    });
  }

  public async getContactLookup(searchResultItem: ISearchResult): Promise<string[]> {
    let contacts: string[] = [];
    const spResponse: SPHttpClientResponse = await this._spHttpClient.get(
        `${this._currentWebUrl}/_api/web/lists('${searchResultItem.IdentityListId}')/items(${searchResultItem.ListItemID})?$select=dmdThemaProductContact/Title&$expand=dmdThemaProductContact`,
        SPHttpClient.configurations.v1
    );
    if (spResponse.ok) {
      const responseData = await spResponse.json();
      console.log(responseData);
      responseData.dmdThemaProductContact.forEach((e: any) => {
        contacts.push(e.Title);
      });
    }

    return contacts;
  }
  
  public async getContactData(title: string): Promise<string> {
    let data: string = "";
    try {

      const spResponse2: SPHttpClientResponse = await this._spHttpClient.get(
        `${this._currentWebUrl}/_api/web/lists/getByTitle('Contactgegevens')/items?$select=dmdContactDetails&$filter=Title eq '${title.replace('&','%26')}'`,
        SPHttpClient.configurations.v1
      );
      if (spResponse2.ok) {
        const responseData = await spResponse2.json(); 
        console.log(responseData);
        data = responseData.value[0].dmdContactDetails;
      }
    } catch (error) {
      console.error(error);
    }
    return data;
  }

}