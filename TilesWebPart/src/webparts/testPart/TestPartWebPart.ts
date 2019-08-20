import { Version, Log } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestPartWebPart.module.scss';
import * as strings from 'TestPartWebPartStrings';

export interface ITestPartWebPartProps {
  description: string;
}

export default class TestPartWebPart extends BaseClientSideWebPart<ITestPartWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.testPart }">
        <div class="${ styles.container }">
          <div id="test" class="${ styles.row }">
            <div class="${ styles.columnleft }">
              <div class="${ styles.boxhalf } ${ styles.box }">
                <span class="${ styles.title }">Welcome to SharePoint!</span>
                <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
                <p class="${ styles.description }">${escape(this.properties.description)}</p>
                <a href="https://aka.ms/spfx" class="${ styles.button }">
                  <span class="${ styles.label }">Learn more</span>
                </a>
              </div>
            </div>
            <div class="${ styles.columnright }">
              <div class="${ styles.boxrow}">
                <div class="${ styles.boxquarter } ${ styles.box }">
                  <span class="${ styles.title }">Welcome to SharePoint!</span>
                  <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
                </div>
                <div class="${ styles.boxquarter } ${ styles.box }">
                  <span class="${ styles.title }">Welcome to SharePoint!</span>
                  <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
                </div>
              </div>
              <div class="${ styles.boxrow}">
                <div class="${ styles.boxquarter } ${ styles.box }">
                  <span class="${ styles.title }">Welcome to SharePoint!</span>
                  <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
                </div>
                <div class="${ styles.boxquarter } ${ styles.box }">
                  <span class="${ styles.title }">Welcome to SharePoint!</span>
                  <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
                </div>
              </div>
            </div>
          </div>

          <div class="${ styles.row } swiper-slide">
            <div class="${ styles.columnleft }">
              <div class="${ styles.boxhalf } ${ styles.box }">
                <span class="${ styles.title }">Welcome to SharePoint!</span>
                <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
                <p class="${ styles.description }">${escape(this.properties.description)}</p>
                <a href="https://aka.ms/spfx" class="${ styles.button }">
                  <span class="${ styles.label }">Learn more</span>
                </a>
              </div>
            </div>
            <div class="${ styles.columnright }">
              <div class="${ styles.boxrow}">
                <div class="${ styles.boxquarter } ${ styles.box }">
                  <span class="${ styles.title }">Welcome to SharePoint!</span>
                  <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
                </div>
                <div class="${ styles.boxquarter } ${ styles.box }">
                  <span class="${ styles.title }">Welcome to SharePoint!</span>
                  <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
                </div>
              </div>
              <div class="${ styles.boxrow}">
                <div class="${ styles.boxquarter } ${ styles.box }">
                  <span class="${ styles.title }">Welcome to SharePoint!</span>
                  <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
                </div>
                <div class="${ styles.boxquarter } ${ styles.box }">
                  <span class="${ styles.title }">Welcome to SharePoint!</span>
                  <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>`;
  }

  public onInit() : Promise<void> {
    return Promise.resolve<void>();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
