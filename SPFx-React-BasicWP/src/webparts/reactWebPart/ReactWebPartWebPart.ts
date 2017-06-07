import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'reactWebPartStrings';

import ToDoContainer from './components/ToDoContainer/ToDoContainer';
import { IToDoContainerProps } from './components/ToDoContainer/IToDoContainerProps';

import { IReactWebPartWebPartProps } from './IReactWebPartWebPartProps';
import {IToDoItem} from './model/IToDoItem';

export default class ReactWebPartWebPart extends BaseClientSideWebPart<IReactWebPartWebPartProps> {
  private mockItems : IToDoItem[] = [{Id : 1, Title : "MockTask 1"}, 
                                                {Id : 2, Title : "MockTask 2"},
                                                {Id : 3, Title : "MockTask 3"},
                                                {Id : 4, Title : "MockTask 4"},];
  public render(): void {
    const element: React.ReactElement<IToDoContainerProps> = React.createElement(
      ToDoContainer,
      {
        description: this.properties.description,
        items : this.mockItems        
      }
    );

    ReactDom.render(element, this.domElement);
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
