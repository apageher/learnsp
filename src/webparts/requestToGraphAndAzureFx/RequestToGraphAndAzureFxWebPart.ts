import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'RequestToGraphAndAzureFxWebPartStrings';
import RequestToGraphAndAzureFx from './components/RequestToGraphAndAzureFx';
import { IRequestToGraphAndAzureFxProps } from './components/IRequestToGraphAndAzureFxProps';

export interface IRequestToGraphAndAzureFxWebPartProps {
  description: string;
}

export default class RequestToGraphAndAzureFxWebPart extends BaseClientSideWebPart<IRequestToGraphAndAzureFxWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IRequestToGraphAndAzureFxProps > = React.createElement(
      RequestToGraphAndAzureFx,
      {
        aadHttpClientFactory: this.context.aadHttpClientFactory,
        msGraphClientFactory: this.context.msGraphClientFactory
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
              groupName: '',
              groupFields: [

              ]
            }
          ]
        }
      ]
    };
  }
}
