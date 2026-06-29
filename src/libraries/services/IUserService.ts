import { IProfileImage } from "../data/IProfileImage";
import { IUserPresence } from "../data/IUserPresence";
import { ISearchResult } from "../data/ISearchResult";

export interface IUserService {
    getUserData(searchResultItem: ISearchResult, expandField: string): Promise<ISearchResult>;
    getImageBase64FromUrl(pictureUrl: string, userId:string): Promise<string>;
    getImageBase64(users: string[]): Promise<IProfileImage>;
    getUserPresence(userObjIds: string[]): Promise<IUserPresence[]>;
    getPresenceStatusCurrent(userId: string): Promise<IUserPresence>;
    getUserId(userProfileProperties: any[]): string;
}