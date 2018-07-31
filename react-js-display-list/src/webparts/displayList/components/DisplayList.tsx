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
import { css } from 'office-ui-fabric-react/lib/Utilities';
//import {HTMLToReact} from 'html-to-react';
export interface dataListItem {
  title?: string;
  created?: any;
  modified?: any;
  aurthor?: any;
  htmlSize?: string;
  htmlLink?: string;
}
export default class DisplayList extends React.Component<IDisplayListProps, IDisplayListState> {
  private _dataProvider: IDataProvider;
  private _dropdownOptions: IPropertyPaneDropdownOption[] = [];
  private _dataListItem: dataListItem;
  private _dataListItems: dataListItem[] = [];
  private element;

  constructor(props: IDisplayListProps) {

    super(props);
    this.setState({
      dropdownOptions: [],
      hello: 'Hello World'
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

        <div className={styles.grid}>
          <div className={styles.row}>
            <div className={css(styles.title, ' ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2')}>Title</div>
            <div className={css(styles.title, ' ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2')}>Created</div>
            <div className={css(styles.title, '  ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2')}>Modified</div>
            <div className={css(styles.title, ' ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2')}>Created By</div>
            <div className={css(styles.title, ' ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2')}>Size</div>
            <div className={css(styles.title, ' ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2')}>Url</div>

          </div>
          <hr />
          <div id="spListContainer"></div>
        </div>
      </div>

    );

  }
  public componentDidMount(): void {

    this._renderAsync();
  }

  private _renderAsync(): void {
    if (Environment.type === EnvironmentType.Local) {
      const title = React.createElement('h1', {}, 'local test environment [No connection to SharePoint]');
      ReactDOM.render(
        title,
        document.getElementById('spListContainer'));
    } else {
      console.log(this.props.dropdownProp);
      if (this.props.dropdownProp === undefined || this.props.dropdownProp == '' || this.props.dropdownProp === null) {
        const title = React.createElement('h1', {}, 'Select proper list name.');
        ReactDOM.render(
          title,
          document.getElementById('spListContainer'));
      }
      else {

        this.props.dataProvider.getListData(this.props.dropdownProp).then((response: ISPLists) => {
          console.log(response.value)
          this._renderList(response.value);
        }).catch((err) => {
          Log.error('js-display-List', err);
          this.context.statusRenderer.renderError(this.props.dataProvider.webPartContext.domElement, err);
        });
      }

    }
  }
  private _createElement(item: dataListItem): any {
    console.log('i am createElement')
    var createdElement = React.createElement('div', { className: styles.row },
      [
        React.createElement('div', { className: `${styles.divCol} ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m` }, item.title),
        React.createElement('div', { className: `${styles.divCol} ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m` }, item.created),
        React.createElement('div', { className: `${styles.divCol} ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m` }, item.modified),
        React.createElement('div', { className: `${styles.divCol} ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m` }, item.aurthor),
        React.createElement('div', { className: `${styles.divCol} ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m` }, item.htmlSize),
        item.htmlLink ?
          React.createElement('div', { className: `${styles.divCol} ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m` },

            React.createElement('a', { href: item.htmlLink }, 'Link')

          ) : React.createElement('div', { className: `${styles.divCol} ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m` }, item.htmlLink)
      ]
    )
    return createdElement;
  }
  private _createElements(items: dataListItem[]): any {
    console.log('i am CreatedElements');
    const initialUserState = {

      arr: []
    }
    items.forEach((item: dataListItem) => {
      var result = this._createElement(item);
      console.log(result);
      initialUserState.arr.push(result);

    })
    console.log(initialUserState);
    return initialUserState.arr;
  }
  public componentDidUpdate(prevDropdownProp): void {

    if (this.props.dropdownProp !== prevDropdownProp) {
      console.log('PreviousElement:', prevDropdownProp);

      this._renderAsync();

    }
  }
  public _renderList(items: ISPList[]): void {
    console.log("I just entered")
    if (!items) {
      console.log("i am if")
      const htmlData = React.createElement('p', { className: 'label' }, "The selected list doesnot exist");
      ReactDOM.render(htmlData, document.getElementById('spListContainer'));

    } else if (items.length === 0) {
      console.log("i am if else")
      const htmlData = React.createElement('p', { className: 'label' }, 'The selected list is empty')
      ReactDOM.render(htmlData, document.getElementById('spListContainer'));

    } else {
      console.log("I am else")
      this._dataListItems = [];
      items.forEach((item: ISPList) => {
        console.log(item);
        let title: string = '';
        let size: string = '';

        let link: string = '';

        if (item.Title === null) {
          console.log("no title:", item.File)
          if (item.File === null && item.Title === null || item.File === undefined && item.Title === null) {
            console.log("no title and file")
            title = "Missing title for item with ID:" + item.Id;
          } else {
            title = item.File.Name;
            size = (item.File.Length / 1024).toFixed(2);
            link = item.File.ServerRelativeUrl;


          }
        }
        else {
          title = item.Title;

        }
        let created: any = item["Created"];
        let modified: any = item["Modified"]
        this._dataListItem = {
          aurthor: item['Author'].Title,
          created: created.substring(0, created.length - 1).replace('T', ' '),
          modified: modified.substring(0, created.length - 1).replace('T', ' '),
          title: title,
          htmlLink: link,
          htmlSize: size


        }
        this._dataListItems.push(this._dataListItem);
        console.log("i am in loop");
      });
      console.log(this._dataListItems)
      const resultElement = this._createElements(this._dataListItems);
      console.log('resulting element:');
      console.log(resultElement);

      const creatingElement = React.createElement('div', {}, resultElement)


      ReactDOM.render(creatingElement, document.getElementById('spListContainer'));

    }

  }
}
