import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import { SPComponentLoader } from '@microsoft/sp-loader';
import * as jQuery from 'jquery';

import { IWebPartProps } from './models';
import ListService from './services/listService';
import DomService from './services/domService';
import { NavigatorService } from './services/navigatorService';
import { ListFieldParseService } from './services/listFieldParseService';
import styles from './TestPartWebPart.module.scss';
import * as strings from 'TestPartWebPartStrings';
import '../../../node_modules/nivo-slider/jquery.nivo.slider.js';
require('../../../node_modules/nivo-slider/nivo-slider.css')
require('./assets/styles.css');
require('./assets/stylesIE.css');

export default class TestPartWebPart extends BaseClientSideWebPart<IWebPartProps> {

  private _navitgatorService: NavigatorService;
  private _listService: ListService;
  private _domService: DomService;
  private _listFieldParseService: ListFieldParseService;
  public _result: any[];
  public _data: any[];
  public _viewName: string;


  private _getAllViewsFromListAjaxCallError: string = '';
  private _loadingExternalJSomScriptResult: any = {
    isSuccess: true,
    message: ''
  };


  public render(): void {
    let self = this;
    self.InjectDependency(function () {
      // success callback
      self.renderWebPartElement();
      self.renderWebPartContent();
    }, function (message: string) {
      // error callback
      self.renderWebPartElement();
      self._domService.displayMessage(message, true);
    });
  }


  // render only elements (slider content will render on next)
  private renderWebPartElement(): void {
    var self = this;
    self.domElement.innerHTML = `<div id="${self._domService._imageRotatorAppHtmlElementId}"></div>`;
    self._domService.renderWebPart(self.properties.IsBrowserIE);
  }

  //render Slider content with help of Rest API call / Jsom
  private renderWebPartContent(): void {
    var self = this;
    // Rest API call to get List Item
    this._listService.getAllListItem().then((data: any[]) => {
      // render slider content from Rest api data (List data)
      self._domService.renderSlideContentFromListData(data);

      if (!self._loadingExternalJSomScriptResult.isSuccess && self._loadingExternalJSomScriptResult.message) {
        let _loadingExternalJsomScriptfailed = `Loading Jsom Script Failed.You can edit loading url from web part properties(edit & reload).`;
        self._domService.displayMessage(_loadingExternalJsomScriptfailed, false);
      }
    }).catch(err => {
      var message: any = err;
      if (message === 404) message = `List '${self.properties.ListName}' does not exist at site with URL '${self.context.pageContext.web.absoluteUrl}'`;
      if (message === 400) message = `No Data : Inappropriate Request (fields or property might wrong)`;
      self._domService.displayMessage(message, true);
    });
    return;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //initialize Jsom script , checking SP present or not  , Decide external jsom script Loading success or not 
  protected onInit(): Promise<void> {

    let self = this;

    this._navitgatorService = new NavigatorService();
    this.properties.IsBrowserIE = this._navitgatorService.isBrowserIE;
    if (!this.properties.JsomScriptBaseURL) this.properties.JsomScriptBaseURL = self.context.pageContext.web.absoluteUrl;

    return super.onInit();
  }


  /* Initialize and inject all dependent service (DomService, ListService , ListFieldParseService)
   Load all available View with selected list to bind in View drop Down*/
  private InjectDependency(succesCallback: any, errorCallback: any): void {

    let self = this;
    if (!self.properties.ListName) self.properties.ListName = 'DataList';
    if (!self.properties.SiteURL) self.properties.SiteURL = self.context.pageContext.web.absoluteUrl;
    if (!self.properties.ViewName) self.properties.ViewName = 'All Items';
    if (!self.properties.SelectedViewName) self.properties.SelectedViewName = self.properties.ViewName;


    // initialize and inject dependent ListFieldParseService class
    self._listFieldParseService = new ListFieldParseService();

    // initialize and inject dependent DomService : which is responssible for all rendering dom element
    self._domService = new DomService(self.properties, self._listFieldParseService, self.context.pageContext.web.absoluteUrl);

    // initialize and inject dependent ListService class with Initialized SP.ClientContext
    self._listService = new ListService(self.properties.SiteURL, self.context.spHttpClient, self.properties, self._domService);


    // load external css
    if (self.properties.externalCssFile && self.properties.externalCssFile !== '') {
      SPComponentLoader.loadCss(self.properties.externalCssFile);
    }

    // on success call back : will render the Dom element and dom content
    succesCallback();
  }

  // set property pane
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Image Rotator is a plugin based on the Nivo Slider plugin for JQuery. This plugin allows you to show Images and link them to SharePoint items"
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description'
                }),
                PropertyPaneTextField('TitleFieldName', {
                  label: 'Title Field '
                }),
                PropertyPaneTextField('TeaserFieldsName', {
                  label: 'Teaser Fields '
                }),
                PropertyPaneTextField('ImageFieldName', {
                  label: 'Image'
                }),
                PropertyPaneTextField('ArticleLinkFieldName', {
                  label: 'Article Link'
                }),
                PropertyPaneTextField('ListName', {
                  label: 'List Name'
                }),
                PropertyPaneTextField('SiteURL', {
                  label: 'Site URL'
                }),
                PropertyPaneTextField('NumberOfitems', {
                  label: 'Number Of Items'
                }),
                PropertyPaneTextField('MaxCharInTeaser', {
                  label: 'Max Char In Teaser'
                }),

                PropertyPaneTextField('RotationSpeed', {
                  label: 'Rotation Speed'
                }),
                PropertyPaneCheckbox('ShowNavigationButton', {
                  checked: this.properties.ShowNavigationButton,  // navigation option change dynamically type : boolean
                  text: 'Show Navigation Button'
                }),
                PropertyPaneTextField('Height', {
                  label: 'Height'
                }),
                PropertyPaneTextField('Width', {
                  label: 'Width'
                }),
                PropertyPaneTextField('externalCssFile', {
                  label: 'External Css'
                }),
                PropertyPaneTextField('JsomScriptBaseURL', {
                  label: 'Jsom Script BaseURL'
                }),
                PropertyPaneCheckbox('EnableResponsive', {
                  checked: this.properties.EnableResponsive,
                  text: 'Enable Responsive'
                }),
                PropertyPaneChoiceGroup('SelectNavigationOptions', {
                  label: "Show Slider Navigation As",
                  options: [
                    { key: 'Thumnails', text: 'Thumnails', checked: true },
                    { key: 'Bullet', text: 'Bullet' },
                    { key: 'Number', text: 'Number' },
                  ],
                })
                //#7 Refactor  : 
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  // on change of Property value
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {

    // when change Navigation option thumbnails, bullet, number
    if (propertyPath === "SelectNavigationOptions") {

      if (newValue === 'Thumnails') this.properties.ShowThumbnails = true;
      else this.properties.ShowThumbnails = false;

      if (newValue === 'Number') {
        this.properties.ShowThumbnails = false;
        this.properties.ShowNavigationNumber = true;
      }
      else {
        this.properties.ShowNavigationNumber = false;
      }
    }

    if (propertyPath === "externalCssFile" && newValue !== '') {
      SPComponentLoader.loadCss(this.properties.externalCssFile);
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }
}

