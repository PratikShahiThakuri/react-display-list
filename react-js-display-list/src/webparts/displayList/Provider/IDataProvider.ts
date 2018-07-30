import { ISPLists } from "../Model/ISPLIsts";
import { ISPList } from "../Model/ISPList";
import { IWebPartContext } from "../../../../node_modules/@microsoft/sp-webpart-base";

export interface IDataProvider{
    webPartContext:IWebPartContext;
    getListTitles():Promise<ISPLists>;
    getListData(listName:string):Promise<ISPLists>;
    //renderList(items:ISPList[]):void;
}