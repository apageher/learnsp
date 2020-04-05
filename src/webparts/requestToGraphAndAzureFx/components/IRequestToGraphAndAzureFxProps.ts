import { AadHttpClientFactory, MSGraphClientFactory } from '@microsoft/sp-http';

export interface IRequestToGraphAndAzureFxProps {
  aadHttpClientFactory: AadHttpClientFactory;
  msGraphClientFactory: MSGraphClientFactory;
}
