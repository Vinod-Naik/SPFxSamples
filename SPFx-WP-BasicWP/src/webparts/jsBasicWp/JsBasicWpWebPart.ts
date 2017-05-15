import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './JsBasicWp.module.scss';
import * as strings from 'jsBasicWpStrings';
import { IJsBasicWpWebPartProps } from './IJsBasicWpWebPartProps';

export default class JsBasicWpWebPart extends BaseClientSideWebPart<IJsBasicWpWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.description)}</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.txtProperty)}</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.toggleProperty.toString())}</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.chkBoxProperty.toString())}</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.sliderProperty.toString())}</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.choiceProperty.toString())}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
                }),
                PropertyPaneTextField('txtProperty',{
                  label: "MultiLine Text",
                  multiline : true
                }),
                PropertyPaneCheckbox('choiceProperty',{
                  text : "Choice Field"
                }),
                PropertyPaneDropdown('choiceProperty',{
                  label:"Dropdown Field",
                  options:[
                    {key:"1",text:"Option1"},
                    {key:"2",text:"Option2"},
                    {key:"3",text:"Option3"},
                    {key:"4",text:"Option4"}
                ]}),
                PropertyPaneSlider('sliderProperty',{
                  label:"Slider Property",
                  max:100,
                  min:0,
                  showValue:true,
                }),
                PropertyPaneToggle('toggleProperty',{
                  label:"Toggle Property",
                  offText:"Off",
                  onText:"On"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
