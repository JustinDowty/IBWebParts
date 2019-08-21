import { Version, Log } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestPartWebPart.module.scss';
import * as strings from 'TestPartWebPartStrings';

import * as $ from 'jquery';
require("slick-carousel");
require("./slick.css");

export interface ITestPartWebPartProps {
  description: string;
  box1BgColor: string;
  box1BgImage: string;
  box1VideoUrl: string;
  box1Url: string;
  box2BgColor: string;
  box2BgImage: string;
  box2VideoUrl: string;
  box2Url: string;
  box3BgColor: string;
  box3BgImage: string;
  box3VideoUrl: string;
  box3Url: string;
  box4BgColor: string;
  box4BgImage: string;
  box4VideoUrl: string;
  box4Url: string;
  box5BgColor: string;
  box5BgImage: string;
  box5VideoUrl: string;
  box5Url: string;
  box6BgColor: string;
  box6BgImage: string;
  box6VideoUrl: string;
  box6Url: string;
  box7BgColor: string;
  box7BgImage: string;
  box7VideoUrl: string;
  box7Url: string;
  box8BgColor: string;
  box8BgImage: string;
  box8VideoUrl: string;
  box8Url: string;
  box9BgColor: string;
  box9BgImage: string;
  box9VideoUrl: string;
  box9Url: string;
  box10BgColor: string;
  box10BgImage: string;
  box10VideoUrl: string;
  box10Url: string;
}

export default class TestPartWebPart extends BaseClientSideWebPart<ITestPartWebPartProps> {
  
  public render(): void {
    this.setTestValues();
    var boxes1 = this.getBoxes(10);
    var boxes2 = this.getBoxes(10);
    var boxes = boxes1.concat(boxes2);
    var body = `<div class="${ styles.testPart }">
                  <div class="${styles.container} slider">`;
    for(let i = 0; i < boxes.length; i+=5) {
      body += `<div class="${ styles.row }">
                  <div class="${ styles.columnleft }">`
                     + this.buildLargeBox(boxes[i]) +
                  `</div>
                  <div class="${ styles.columnright }">
                      <div class="${ styles.boxrow}">`
                        + this.buildSmallBox(boxes[i+1]) + this.buildSmallBox(boxes[i+2]) +
                      `</div>
                      <div class="${ styles.boxrow}">`
                      + this.buildSmallBox(boxes[i+3]) + this.buildSmallBox(boxes[i+4]) +
                    `</div>
                  </div>
                </div>`
    }
    body += `</div></div>`;
    
    this.domElement.innerHTML = `
      <div class="${ styles.testPart }">
        <div class="${styles.container} slider">`
          + body +
        `</div>
      </div>`;
  
      ($ as any)('.slider').slick({
        arrows: false,
        dots: true
      });
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

  private getBoxes(num){
    var boxes = [];
    for(let i = 1; i <= num; i++) {
      let color = "box" + i + "BgColor";
      let image = "box" + i + "BgImage";
      let video = "box" + i + "VideoUrl";
      let url = "box" + i + "Url";
      let item = {
        color: this.properties[color],
        image: this.properties[image],
        video: this.properties[video],
        url: "www.google.com",
        title: "Block Main Title",
        subTitle: "I am a subtitle!"
      }
      boxes.push(item);
    }
    return boxes;
  }

  private buildLargeBox(box) {
    var largeBox = box.video != ''
      ? `<div class="${styles.largeVidWrapper}"><iframe class="${styles.video}" src="${box.video}" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></div>` 
      : `<a href="www.google.com" class="${ styles.box }">
          <div class="${styles.content}">
            <div class="${ styles.title }">${box.title}</div>
            <div class="${ styles.subTitle }">${box.subTitle}</div>
          </div>
          <div style="background: ${box.color} url(${box.image})" class="${ styles.bg}"></div>
        </a>`;
    return largeBox;
  }

  private buildSmallBox(box) {
    var smallBox = box.video != ''
      ? `<div class="$ ${ styles.box }">
          <iframe class="${styles.video}"  src="${box.video}" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>
        </div>`
      : `<a href="www.google.com" class="${ styles.box }">
          <div class="${styles.content}">
            <div class="${ styles.smallTitle }">Welcome to SharePoint!</div>
            <div class="${ styles.smallSubTitle }">${box.subTitle}</div>
          </div>
          <div style="background: ${box.color} url(${box.image})" class="${ styles.bg}"></div>
        </a>`
    return smallBox;
  }

  private setTestValues(){
    this.properties.box1BgColor = '#172a54';
    this.properties.box1BgImage = "http://fc02.deviantart.net/fs34/i/2008/304/3/d/Triforce_2_by_5995260108.png";
    this.properties.box1VideoUrl = "";
    this.properties.box2BgColor = '#172a54';
    this.properties.box2BgImage = "#a8353a";
    this.properties.box2VideoUrl = "https://www.youtube.com/embed/uAOR6ib95kQ";
    this.properties.box3BgColor = '#172a54';
    this.properties.box3BgImage = "http://fc02.deviantart.net/fs34/i/2008/304/3/d/Triforce_2_by_5995260108.png";
    this.properties.box3VideoUrl = "";
    this.properties.box4BgColor = "#a8353a";
    this.properties.box4BgImage = "#a8353a";
    this.properties.box4VideoUrl = "";
    this.properties.box5BgColor = '#172a54';
    this.properties.box5BgImage = "";
    this.properties.box5VideoUrl = "";
    this.properties.box6BgColor = '#172a54';
    this.properties.box6BgImage  = "http://fc02.deviantart.net/fs34/i/2008/304/3/d/Triforce_2_by_5995260108.png";
    this.properties.box6VideoUrl = "https://www.youtube.com/embed/uAOR6ib95kQ";
    this.properties.box7BgColor = '#a8353a';
    this.properties.box7BgImage = "";
    this.properties.box7VideoUrl = "";
    this.properties.box8BgColor = '#172a54';
    this.properties.box8BgImage = '';
    this.properties.box8VideoUrl = "";
    this.properties.box9BgColor = '#172a54';
    this.properties.box9BgImage = "";
    this.properties.box9VideoUrl = "";
    this.properties.box10BgColor = '#172a54';
    this.properties.box10BgImage = "";
    this.properties.box10VideoUrl = "https://www.youtube.com/embed/uAOR6ib95kQ";
  }
}
