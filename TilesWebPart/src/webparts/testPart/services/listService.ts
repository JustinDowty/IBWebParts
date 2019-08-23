import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { UrlService } from "./urlService";
import { IWebPartProps } from "../models";
import DomService from '../services/domService';

export default class ListService {

    private _siteUrl: string;
    private _spHttpClient: SPHttpClient;
    private _webPartProps: IWebPartProps;

    constructor(
        siteUrl: string,
        spHttpClient: SPHttpClient,
        webPartProps: IWebPartProps,
        private domService: DomService) {
        this._siteUrl = siteUrl;
        this._spHttpClient = spHttpClient;
        this._webPartProps = webPartProps;
    }


    public getAllListItem(): Promise<any[]> {

        let self = this;
        let promise = new Promise<any[]>((resolve, reject) => {

            var col_Title = encodeURI(self._webPartProps.TitleFieldName);
            var col_Image = encodeURI(self._webPartProps.ImageFieldName);
            var col_ArticleLink = encodeURI(self._webPartProps.ArticleLinkFieldName);
            //#3 Refactor : Columns Need to Fetch From The List
            let columns = `${col_Title},${col_Image},${col_ArticleLink}`;
            let requestUrl = UrlService.BuildRequestUrl(self._siteUrl, self._webPartProps.ListName, columns, self._webPartProps.NumberOfitems);
            self._spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
                .then((response: SPHttpClientResponse) => {
                    if (response.ok) {
                        response.json().then((responseJSON) => {
                            if (responseJSON != null && responseJSON.value != null) {
                                let items: any[] = responseJSON.value;
                                resolve(items);
                            } else {
                                resolve(null);
                            }
                        });
                    } else {
                        reject(response.status);
                    }
                }).catch(err => {
                    reject(err);
                });
        });

        return promise;
    }

    // After getting Jsom call Data , iterate and parse each item into object to process
    private parseCurrentItemIntoListForRender(currentItem: any): any {

        let self = this;
        var item = {};
        var _Title: string = '';
        var _Teaser: string = '';
        var _Image: string = '';
        var _ArticleLink: string = ''; 

        var _ID = currentItem.get_id();
        var contentTypeID = currentItem.get_item("ContentTypeId");
        var col_Title = encodeURI(self._webPartProps.TitleFieldName);
        var col_Teaser = encodeURI(self._webPartProps.TeaserFieldsName);
        var col_Image = encodeURI(self._webPartProps.ImageFieldName);
        var col_ArticleLink = encodeURI(self._webPartProps.ArticleLinkFieldName);

        // handled : if the given Title /Teaser Field name wrong/not present in list or views items
        try {
            _Title = currentItem.get_item(col_Title);
            _Teaser = currentItem.get_item(col_Teaser);
            _Image = currentItem.get_item(col_Image);
            _ArticleLink = currentItem.get_item(col_ArticleLink);


            item[self._webPartProps.TitleFieldName] = _Title;
            item[self._webPartProps.TeaserFieldsName] = _Teaser;
            item[self._webPartProps.ImageFieldName] = _Image;
            item[self._webPartProps.ArticleLinkFieldName] = _ArticleLink;
            item['ID'] = _ID;

        } catch (err) {
            var message: any = err;
             message = `Fileds unmatched : Not appropiate fields`;
            self.domService.displayMessage(message, true);
            return false;
  
        }





        // determin whether item is from wiki page or publishing page
        if (currentItem) item['ContentType'] = self.determineListItemIsWikiPageOrPublishingPage(currentItem);
        if (currentItem) item['FileLeafRef'] = currentItem.get_item('FileLeafRef');

        // It's possible to create Duplicate Title column in Site Pages (List) in share point .Handled To get right Title in such case
        try {

            if (item['ContentType'] && self._webPartProps.TitleFieldName == 'Title' && (item['ContentType'] === self.domService._determineContentTypePage.wikiPage || item['ContentType'] === self.domService._determineContentTypePage.publishingPage)) {
                if (currentItem.get_item('Title0') != null) {
                    let _title = currentItem.get_item('Title0');
                    item[self._webPartProps.TitleFieldName] = _title;
                }
                else {
                    item[self._webPartProps.TitleFieldName] = '';
                }
            }
        } catch (err) {

            item[self._webPartProps.TitleFieldName] = _Title;
        }


        return item;
    }

    // determine current item is a wiki page or publishing modern page or web part page
    public determineListItemIsWikiPageOrPublishingPage(currentItem: any): string {

        var _pageType: string = 'Not Determined';
        var contentTypeID = currentItem.get_item("ContentTypeId");
        // var currentItemWikiField = currentItem.get_item("WikiField");

        if (contentTypeID.toString().startsWith("0x010108"))
            _pageType = 'WikiPage';
        else if (contentTypeID.toString().startsWith("0x010109"))
            _pageType = 'WebPartPage';
        else if (contentTypeID.toString().startsWith("0x010100")) {
            _pageType = 'publishingPage';
        }

        return _pageType;
    }

    public get baseUrl(): string {
        return this._siteUrl;
    }
}