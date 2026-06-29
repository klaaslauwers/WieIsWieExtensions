import { ISearchResult } from "../data/ISearchResult";

export interface ISharePointService {
    getContactData(title: string): Promise<string>;
    getContactLookup(searchResultItem: ISearchResult): Promise<string[]>;
}