import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'DisplayListWebPartStrings';
import DisplayList from './components/DisplayList';
import { IDisplayListProps } from './components/IDisplayListProps';
import { Service } from './Provider/Service';
import { ISPList } from './Model/ISPList';
import { IDataProvider } from './Provider/IDataProvider';
import * as ReactDOM from 'react-dom';

export interface IDisplayListWebPartProps {
  dropdownProp: string;
  dataProvider:IDataProvider;

}

export default class DisplayListWebPart extends BaseClientSideWebPart<IDisplayListWebPartProps> {

  private _dataProvider :IDataProvider;
  
  private _dropdownOptions: IPropertyPaneDropdownOption[] = []; 
  private _displayListComponent: DisplayList;
  
protected onInit():Promise<void>{
   
  // if (DEBUG && Environment.type === EnvironmentType.Local) {
  //   const title = React.createElement('p', {}, 'local test environment [No connection to SharePoint]');
  //     ReactDOM.render(
  //       title,
  //       document.getElementById('spListContainer'));

  // } 
  // else {
    
      this._dataProvider = new Service();
      this._dataProvider.webPartContext = this.context;
      this._getListsName();
      // }
 
    return super.onInit();
  }
public render(): void {
    
  // this._dataProvider = new Service();
  // this._dataProvider.webPartContext = this.context;
    const element: React.ReactElement<IDisplayListProps > = React.createElement(
      DisplayList,
      {
        dropdownProp: this.properties.dropdownProp,
        dataProvider:this._dataProvider
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  private _getListsName():Promise<any>{
    if(DEBUG && Environment.type===EnvironmentType.Local){
      // const title = React.createElement('p', {}, 'local test environment [No connection to SharePoint]');
      //   ReactDOM.render(
      //    title,
      //    document.getElementById('spListContainer'));
    }else{
   
    this._dataProvider.getListTitles().then((response)=>{
      
      this._dropdownOptions=response.value.map((list:ISPList)=>{
        
        return{
          key:list.Title,
          text:list.Title
        };
      });
    });
    }
    return Promise.resolve();
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
                PropertyPaneDropdown('dropdownProp', {
                  label: 'List Title',
                  options:this._dropdownOptions
                })
              ]
            }
          ]
        }
      ]
    };
  }
 
}
