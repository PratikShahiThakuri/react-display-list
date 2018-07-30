import { IDataProvider } from "./IDataProvider";
import { WebPartContext, IWebPartContext } from "@microsoft/sp-webpart-base";
import { ISPList } from "../Model/ISPList";
import { ISPLists } from "../Model/ISPLIsts";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { Log } from "@microsoft/sp-core-library";

import * as ReactDOM from "react-dom";

export class Service implements IDataProvider{
   
    
    private _webPartContext:IWebPartContext;
    private _listsUrl:string;
    private _ispLists:ISPLists;
    public set webPartContext(value: IWebPartContext) {
        this._webPartContext = value;
        this._listsUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists`;
      }
    
      public get webPartContext(): IWebPartContext {
        return this._webPartContext;
      }
      public getListTitles():Promise<ISPLists>{
        const queryString:string='?$filter=Hidden eq false';
        const queryUrl: string = this._listsUrl + queryString;
          return this._webPartContext.spHttpClient.get(queryUrl,SPHttpClient.configurations.v1)
          .then((response:SPHttpClientResponse)=>{
              return response.json();
          });
          
      }
      public getListData(listName:string):Promise<ISPLists>{
        const queryString:string ='$select=Title,ID,Created,Modified,Author/ID,Author/Title,File&$expand=Author/ID,Author/Title,File';
        const queryUrl: string = this._listsUrl+`/GetByTitle('${listName}')/items?`+ queryString;
        return this._webPartContext.spHttpClient.get(queryUrl,SPHttpClient.configurations.v1)
        .then((response:SPHttpClientResponse)=>{
            if(response.status===404){
                 Log.error('js-display-List',new Error('List Not Found.'));
                return[];
              }else{
                
                return response.json();
              }
        })
      }
     
    
}