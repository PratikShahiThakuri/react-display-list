import * as React from 'react';
import styles from './DisplayList.module.scss';
import { IDisplayListProps } from './IDisplayListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IDisplayListState } from './IDisplayListState';
import { Environment, EnvironmentType, Log } from '@microsoft/sp-core-library';
import { IDataProvider } from '../Provider/IDataProvider';
import { Service } from '../Provider/Service';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-webpart-base';
import { ISPList } from '../Model/ISPList';
import * as ReactDOM from 'react-dom';
import { ISPLists } from '../Model/ISPLIsts';
//import {HTMLToReact} from 'html-to-react';
export interface dataListItem{
  title?:string;
  created?:any;
  modified?:any;
  aurthor?:any;
  htmlSize?:string;
  htmlLink?:string;
}
export default class DisplayList extends React.Component<IDisplayListProps,IDisplayListState> {
  private _dataProvider :IDataProvider;
  private _dropdownOptions: IPropertyPaneDropdownOption[] = []; 
  private _dataListItem:dataListItem;
  private _dataListItems:dataListItem[]=[];
  constructor(props: IDisplayListProps) {
    
    super(props);
   this.setState({
     dropdownOptions:[],
     hello:'Hello World'
   });
   
    
  }
  public render(): React.ReactElement<IDisplayListProps> {
    

    return (
       <div className={styles.displayList}>
          <p className={styles.para}>
          <span className={styles.label}>
            {this.props.dropdownProp}
              </span>
              
          </p>
        
          <div>
             <div className={styles.row}>
                <div className={styles.divCol}>Title</div>
                <div className={styles.divCol}>Created</div>
                <div className={styles.divCol}>Modified</div>
                <div className={styles.divCol}>Created By</div>
                <div className={styles.divCol}>Size</div>
                <div className={styles.divCol}>Url</div>

              </div>
              <hr />
              <div id="spListContainer"></div>
        </div>
        </div>
        
    );
   
  }
  public componentDidMount():void {

   this._renderAsync();
  }
  private _renderAsync():void{
    if(Environment.type===EnvironmentType.Local){
      const title = React.createElement('h1', {}, 'local test environment [No connection to SharePoint]');
      ReactDOM.render(
        title,
        document.getElementById('spListContainer'));
    }else{
      console.log(this.props.dropdownProp);
      if(this.props.dropdownProp===undefined|| this.props.dropdownProp==''||this.props.dropdownProp===null){
        const title = React.createElement('h1', {}, 'Select proper list name.');
        ReactDOM.render(
        title,
        document.getElementById('spListContainer'));
        }
      else
      {
      
        this.props.dataProvider.getListData(this.props.dropdownProp).then((response:ISPLists)=>{
          console.log(response.value)
         this._renderList(response.value);
        }).catch((err) => {
         Log.error('js-display-List', err);
         this.context.statusRenderer.renderError(this.props.dataProvider.webPartContext.domElement, err);
       });
      }
    
    }
  }
  public componentDidUpdate(prevDropdownProp):void{
    if(this.props.dropdownProp!==prevDropdownProp){
      this._renderAsync();

    }
  }
  public _renderList(items: ISPList[]): void {
    console.log("I just entered")
    if(!items){
      console.log("i am if")
            const htmlData = React.createElement('p', {className:'label'},"The selected list doesnot exist");
            ReactDOM.render(htmlData,document.getElementById('spListContainer'));
            //html='<br/><p class="ms-font-m-plus">The selected list doesnot exist.</p>';
             }else if(items.length===0){
               console.log("i am if else")
               const htmlData = React.createElement('p',{className:'label'},'The selected list is empty')
               ReactDOM.render(htmlData,document.getElementById('spListContainer'));
                     //html='<br/><p class="ms-font-m-plus"> The selected list is empty</p>';
                     }else{
                       console.log("I am else")
             items.forEach((item:ISPList)=>{
     console.log(item);
             let title :string='';
            let size :string='';
            //let sizeHtml:string='';
         let link : string='';
         let linkhtml:string='';
         if(item.Title===null){
             if(item.File===null && item.Title===null){
                 title="Missing title for item with ID= "+ item.Id;
                 }else {
                title=item.File.Name;
            size = (item.File.Length/1024).toFixed(2);
                link = item.File.ServerRelativeUrl;
               const sizeHtml=`<div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
      ${size}
    </div>`
    linkhtml =`<div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
    <a href="${link}">Link</a>
  </div>`
    }
    }
    else{
      title=item.Title;
    }
    let created:any =item["Created"];
    let modified :any =item["Modified"]
  //   html+=`
  //   <div class =" ms-Grid-row ">
  //   <div class=" ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
  //   ${title}               
  //   </div>
  //   <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
  //   ${created.substring(0,created.length -1).replace('T',' ')}
  //   </div>
  //   <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
  //   ${modified.substring(0,created.length -1).replace('T',' ')}
  //   </div>
  //   <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
  //   ${item['Author'].Title}
  //   </div>
  //  ${sizeHtml}
  //   ${linkhtml}
    
  //   </div>
    
    
  //   `
  this._dataListItem={
    aurthor:item['Aurthor'].Title,
    created:created.substring(0,created.length -1).replace('T',' '),
    modified:modified.substring(0,created.length -1).replace('T',' '),
    title:title,
    htmlLink:link,
    htmlSize:size


  }
  this._dataListItems.push(this._dataListItem);
console.log("i am in loop");
  });
  console.log(this._dataListItems)
}

  }
}
