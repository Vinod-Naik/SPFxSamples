import * as React from 'react';
import styles from './ReactWebPart.module.scss';
import { IReactWebPartProps } from './IReactWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class ReactWebPart extends React.Component<IReactWebPartProps, void> {
  public render(): React.ReactElement<IReactWebPartProps> {
    return (
      <div>
        <span>Welcome to SharePoint!</span>
        <p>Customize SharePoint experiences using Web Parts.</p>
        <p>{escape(this.props.description)}</p>
      </div>
    );
  }
}
