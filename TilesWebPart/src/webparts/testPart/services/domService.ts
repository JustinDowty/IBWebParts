import * as jQuery from 'jquery';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IWebPartProps } from '../models';


// styles import
import styles from '../TestPartWebPart.module.scss';
import { ListFieldParseService } from './listFieldParseService';
import { UrlService } from "./urlService";

// import fortawesome
import fontawesome from '@fortawesome/fontawesome'
import faFreeSolid, { faAddressBook } from '@fortawesome/fontawesome-free-solid'


export default class DomService {

    // Id and classes repeatedly used in dom rendering
    public _mainAppContainerHtmlElementId: string = 'mainAppContainer';
    public _imageRotatorAppHtmlElementId: string = 'imageRotatorApp';
    private _sliderWrapperHtmlElementId: string = 'wrapper';
    private _sliderHtmlelementId: string = 'slider';
    private _sliderWrapperNivoClass: string = 'slider-wrapper';
    private _nivoThemeDefaultClass: string = 'theme-default';
    private _numberNavigationClass: string = 'nivo-navigation-number';
    private _bulletNavigationClass: string = 'nivo-navigation-bullet';
    private _thumbnailsNavigationClass: string = 'nivo-navigation-thumbnails';


    private _properties: IWebPartProps;
    private _baseUrl: string;

    // value for Content Type (List item's content type)
    public _determineContentTypePage: any = {
        wikiPage: 'WikiPage',
        publishingPage: 'publishingPage',
        webPartPage: 'WebPartPage',
        notDetermined: 'Not Determined'
    };


    constructor(
        private properties: IWebPartProps,
        private listFieldParseService: ListFieldParseService,
        private baseUrl: string) {

        this._properties = properties;
        this._baseUrl = baseUrl;
    }



    // render initial dom element like wrapper , slider wrapper to ready for nivo slider content to be rendered
    renderWebPart(isBrowserIE: boolean): void {

        // isBrowserIE says true : IE browser , false : any other broser
        let $ = jQuery;
        let self = this;
        var sliderWrapperElementInlineStyles = `height: ${self.properties.Height}px; width: ${self.properties.Width}px;`;
        if (self.properties.EnableResponsive) sliderWrapperElementInlineStyles = `max-height: ${self.properties.Height}px; max-width: ${self.properties.Width}px;`;

        var enableResponsiveStyles = `height: ${self.properties.Height}px; width: ${self.properties.Width}px;`;
        if (self.properties.EnableResponsive) enableResponsiveStyles = `max-height: ${self.properties.Height}px; max-width: ${self.properties.Width}px;`;


        var isIECssClass = '';
        if (isBrowserIE) isIECssClass = 'IE';

        if (isBrowserIE) enableResponsiveStyles = enableResponsiveStyles + `height: ${self.properties.Height}px; width: ${self.properties.Width}px;`;

        let $appDescriptionElement = $(`<div id="${self._mainAppContainerHtmlElementId}" class="${styles.container}"> ${self.properties.description}</div>`);
        $(`#${self._imageRotatorAppHtmlElementId}`).append($appDescriptionElement);

        let $sliderWrapperElement = $(`<div id="wrapper" style="${sliderWrapperElementInlineStyles}"> </div>`);
        $(`#${self._imageRotatorAppHtmlElementId}`).append($sliderWrapperElement);

        let $nivoSliderContainerHtml = `<div class="slider-wrapper theme-default"><div id="${self._sliderHtmlelementId}" style="${enableResponsiveStyles}" class="${isIECssClass}"></div></div>`
        $sliderWrapperElement.html($nivoSliderContainerHtml);

        let _icon = fontawesome.icon(faAddressBook);
    }

    // Render Nivo Slider Content : After getting data from REST API call to appropiate List
    renderSlideContentFromListData(data: any[]): void {

        let $ = jQuery;
        let self = this;
        var counter: number = 0;
        data.forEach(function (item) {

            // process data item into title / teaser / ArticleLink /Caption ready for render
            // will return how many item processed (fit to show in slider check also has image url and title /not)

            counter = self.processItemToNivoSliderContentForRender(item, counter);
        });

        if (counter == 0) {

            // display a message in slider if there are no data to show (or eligable to show *img url/ *title )
            self.displayMessage('No data found to show', true);
            return;
        }






        // render the main body of nivo slider with subsequent element and content
        self.renderNivoSliderMainBody();

        // manage css class for navigation type (bullet/ thumbnails/number)
        self.manageNivoSliderNavigationClass();

    }


    processItemToNivoSliderContentForRender(item: any, counter: number): number {

        let self = this;


        // encode properties
        var col_Title = encodeURI(self._properties.TitleFieldName);
        var col_Teaser = encodeURI(self._properties.TeaserFieldsName);
        var col_Image = encodeURI(self._properties.ImageFieldName);
        var col_ArticleLink = encodeURI(self._properties.ArticleLinkFieldName);

        var col_ListName = encodeURI(self._properties.ListName);
        var _listName = self._properties.ListName.split(' ').join('');

        // get  title /imgsrc/teaser /article link out from items 
        let imgSrc = item[col_Image];
        let title: string = item[col_Title];
        let teaser = item[col_Teaser];
        let articleLink = item[col_ArticleLink];
        let itemID = item['ID'];

        // check it has any content type (usually have when it come from any wiki /publishind/web part pages)
        // if not then set it to 'Not Determined' when it comes from List itself
        if (!item['ContentType']) item['ContentType'] = self._determineContentTypePage.notDetermined;

        // build blank ArticleLink base on from where it come List Items/ Wiki Pages/Publishing pages
        if (!articleLink && item['ContentType']) articleLink = self.buildArticleLinkForBlankItem(item['FileLeafRef'], item['ContentType'], self._baseUrl, _listName, itemID);


        //restrict if teaser is more than maximum range
        if (teaser == null) {
            teaser = '';
        }
        else {
            if (teaser.length > self._properties.MaxCharInTeaser) {
                teaser = teaser.substring(0, self._properties.MaxCharInTeaser - 3) + "...";
            }

        }


        // if title and image is not blank then process and render image content/ caption
        if (title && imgSrc) {

            let nivoCaptionId = `nivocaptionId-${counter}`;
            let sliderItem = $(`<a href="${articleLink}" target="_blank" />`);

            // Appened and render Based on Browser IE or not
            self.appendImageContentBasedOnBrowser(imgSrc, nivoCaptionId, sliderItem);

            let _slideTitleAndTeaserHtml = `<div id="${nivoCaptionId}" class="nivo-caption">     
                                            <p class="nivo-caption-title"><strong>${title}</strong></p> 
                                            <p class="nivo-caption-teaser">${teaser}</p> 
                                            </div>`;

            $(`#${self._sliderWrapperHtmlElementId}`).append(_slideTitleAndTeaserHtml);

            counter++;

        }

        return counter;
    }



    buildArticleLinkForBlankItem(fieldRef: string, contentyType: string, baseurl: string, listname: string, itemID: string): string {

        let self = this;
        var _articleLinkBasedOnContentType = '';
        _articleLinkBasedOnContentType = UrlService.BuildListItemUrl(baseurl, listname, itemID);
        return _articleLinkBasedOnContentType;
    }


    // manage css class for navigation type (bullet/ thumbnails/number)
    manageNivoSliderNavigationClass(): void {

        let self = this;

        //if want to show navigation in number , need to remove the theme class applied for nivo slider
        if (self.properties.ShowNavigationNumber && $(`.${self._sliderWrapperNivoClass}`).hasClass(self._nivoThemeDefaultClass)) {
            $(`.${self._sliderWrapperNivoClass}`).removeClass(self._nivoThemeDefaultClass);
        }
        //if want to show navigation in Thumbnails or Bullent , need to add/available the theme class applied for nivo slider
        else if (!self.properties.ShowNavigationNumber && !$(`.${self._sliderWrapperNivoClass}`).hasClass(self._nivoThemeDefaultClass)) {
            $(`.${self._sliderWrapperNivoClass}`).addClass(self._nivoThemeDefaultClass);

        }

        // add dedicated css class for Navigation : Number (get rid of other classes)
        if (self.properties.ShowNavigationNumber && !$(`.${self._sliderWrapperNivoClass}`).hasClass(self._numberNavigationClass)) {
            $(`.${self._sliderWrapperNivoClass}`).find('.nivo-controlNav').removeClass(self._bulletNavigationClass);
            $(`.${self._sliderWrapperNivoClass}`).find('.nivo-controlNav').removeClass(self._thumbnailsNavigationClass);
            $(`.${self._sliderWrapperNivoClass}`).find('.nivo-controlNav').addClass(self._numberNavigationClass);
        }

        // add dedicated css class for Navigation : Thumbnails (get rid of other classes)
        if (self._properties.ShowThumbnails && !self.properties.ShowNavigationNumber) {
            $(`.${self._sliderWrapperNivoClass}`).find('.nivo-controlNav').removeClass(self._numberNavigationClass);
            $(`.${self._sliderWrapperNivoClass}`).find('.nivo-controlNav').removeClass(self._bulletNavigationClass);
            $(`.${self._sliderWrapperNivoClass}`).find('.nivo-controlNav').addClass(self._thumbnailsNavigationClass); //nivo-controlNav
        }

        // add dedicated css class for Navigation : Bullet (get rid of other classes)
        if (self._properties.ShowNavigationButton && !self._properties.ShowThumbnails && !self.properties.ShowNavigationNumber) {
            $(`.${self._sliderWrapperNivoClass}`).find('.nivo-controlNav').removeClass(self._numberNavigationClass);
            $(`.${self._sliderWrapperNivoClass}`).find('.nivo-controlNav').removeClass(self._thumbnailsNavigationClass);
            $(`.${self._sliderWrapperNivoClass}`).find('.nivo-controlNav').addClass(self._bulletNavigationClass);
        }
    }

    renderNivoSliderMainBody(): void {

        let $ = jQuery;
        let self = this;

        // Show on slider as content

        (<any>$(`#${self._sliderHtmlelementId}`)).nivoSlider({

            effect: 'random',                                           // Specify sets like: 'fold,fade,sliceDown' 
            slices: 15,                                                 // For slice animations 
            boxCols: 8,                                                 // For box animations 
            boxRows: 4,                                                 // For box animations 
            animSpeed: 1000,                                            // Slide transition speed 
            pauseTime: self._properties.RotationSpeed,                  // How long each slide will show 
            startSlide: 0,                                              // Set starting Slide (0 index) 
            directionNav: self._properties.ShowNavigationButton,        // Next & Prev navigation 
            controlNav: true,                                           // 1,2,3... navigation 
            controlNavThumbs: self._properties.ShowThumbnails,
            // Use thumbnails for Control Nav 
            pauseOnHover: true,                                         // Stop animation while hovering 
            manualAdvance: false,                                       // Force manual transitions 
            prevText: '<i class="fas fa-arrow-alt-circle-left"></i>',   //#6 Refactor : use Icon Font Awsome 
            nextText: '<i class="fas fa-arrow-alt-circle-right"></i>',  //#6 Refactor : use Icon Font Awsome 
            randomStart: false,                                         // Start on a random slide 
            beforeChange: function () { },                              // Triggers before a slide transition 
            afterChange: function () { },                               // Triggers after a slide transition 
            slideshowEnd: function () { },                              // Triggers after all slides have been shown 
            lastSlide: function () { },                                 // Triggers when last slide is shown 
            afterLoad: function () { }
        });
    }

    // Appened and render Based on Browser IE or not
    appendImageContentBasedOnBrowser(imgSrc: string, nivoCaptionId: string, sliderItem: any): void {

        let self = this;

        if (self._properties.IsBrowserIE) {
            let sliderItemImage = $(`<img class="slider-item-image" src="${imgSrc}" data-thumb="${imgSrc}" title="#${nivoCaptionId}"  />`);
            let sliderItemImageWrapper = $(`<div class="slider-item-image-wrapper" style="background-image: url('${imgSrc}')" ></div>`);

            sliderItem.append(sliderItemImageWrapper);
            sliderItem.append(sliderItemImage);

            $(`#${self._sliderHtmlelementId}`).append(sliderItem);
        }
        else {
            let sliderItemImage = $(`<img class="target" src="${imgSrc}" data-thumb="${imgSrc}" title="#${nivoCaptionId}"  />`);
            sliderItem.append(sliderItemImage);

            $(`#${self._sliderHtmlelementId}`).append(sliderItem);
        }
    }


    // display message for Error or handle No Data
    displayMessage(message: string, closeSlider: boolean): void {
        let $ = jQuery;
        var $messageContainerElement: any;

        var _showErrorElementId: string = '';
        var _elementToAppendMessageConatainer: string = '';

        // Either Error message show Along with Slider itself/ Or with out slider
        // Not show the slider , only error message
        if (closeSlider) {

            _showErrorElementId = "showErrorWithOutSlider";

            //  if previous  error message already , remove it to avoid duplicacy 
            $(`#${_showErrorElementId}`).remove();

            _elementToAppendMessageConatainer = this._imageRotatorAppHtmlElementId;
            $messageContainerElement = $(`<div id="${_showErrorElementId}" ><div ><p style="margin-right: 30px;"><strong>${message}</strong></p></div></div>`);

            $(`#${this._sliderWrapperHtmlElementId}`).remove();
        }
        // show the slider with content and also error message
        else {
            _showErrorElementId = "showErrorWithSlider";
            //  if previous  error message already , remove it to avoid duplicacy 
            $(`#${_showErrorElementId}`).remove();

            _elementToAppendMessageConatainer = this._mainAppContainerHtmlElementId;
            $messageContainerElement = $(`<div id="${_showErrorElementId}" ><div ><p style="margin-right: 30px;"><strong>${message}</strong></p><span id ="errorMessage"  class="closeMessageWithSlider">&times;</span></div></div>`);
        }



        $(`#${_elementToAppendMessageConatainer}`).append($messageContainerElement);

        var span = $('#errorMessage');
        var showErrorWithSliderdiv = $(`#${_showErrorElementId}`);

        //Calls the function onclick
        span.on("click", function () {
            showErrorWithSliderdiv.remove();
        })

    }


}