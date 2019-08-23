
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneCheckbox,
    PropertyPaneDropdown,
    IPropertyPaneDropdownOption,
    IPropertyPaneChoiceGroupOption,
    IPropertyPaneChoiceGroupProps
} from '@microsoft/sp-webpart-base';

export interface IWebPartProps {
    description: string;
    TitleFieldName: string;
    TeaserFieldsName: string;
    ImageFieldName: string;
    ArticleLinkFieldName: string;
    ListName: string;
    SiteURL: string,
    ViewName: string,
    NumberOfitems: number;
    MaxCharInTeaser: number;
    RotationSpeed: string;
    ShowNavigationButton: boolean;
    Height: string;
    Width: string;
    externalCssFile: string;
    EnableResponsive: boolean;
    AllAvailableViews: IPropertyPaneDropdownOption[];
    SelectedViewName: string;
    ShowThumbnails: boolean;
    SelectNavigationOptions: string[];
    ShowNavigationNumber: boolean;
    IsBrowserIE: boolean;
    JsomScriptBaseURL: string;
}

