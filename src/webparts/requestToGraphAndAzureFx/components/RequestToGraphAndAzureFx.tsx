import * as React from 'react';
import styles from './RequestToGraphAndAzureFx.module.scss';
import { IRequestToGraphAndAzureFxProps } from './IRequestToGraphAndAzureFxProps';
import { AadHttpClient } from '@microsoft/sp-http';
import { PrimaryButton } from 'office-ui-fabric-react';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IRequestToGraphAndAzureFxState {
  returnAzureFx: string;
}

export default class RequestToGraphAndAzureFx extends React.Component<IRequestToGraphAndAzureFxProps, IRequestToGraphAndAzureFxState> {

  constructor(props: IRequestToGraphAndAzureFxProps) {
    super(props);
    this.state = {
      returnAzureFx: ''
    };
  }

  public clickme = async () => {
    try {
      //Graph API
      const clientGraph = await this.props.msGraphClientFactory.getClient();
      const graphData: MicrosoftGraph.User = await clientGraph.api("/me").get();
      console.log(graphData.displayName);

      //Azure function securizada
      const client = await this.props.aadHttpClientFactory.getClient("https://peichfunction.azurewebsites.net");//Cliente autenticado
      //const client = await this.props.aadHttpClientFactory.getClient("f1a1bc50-fa40-4655-a6b6-7d2cad2607f3"); //También sirve
      //es el Id. de aplicación (cliente) de nuestra Azure function que aparece en el Azure Active directory
      const response = await client.get(`https://peichfunction.azurewebsites.net/api/HttpTrigger1?name=${graphData.displayName}`, AadHttpClient.configurations.v1);
      const data = await response.text(); //response.json() si la Azure function devuelve un Json
      console.log(data);
      this.setState({
        returnAzureFx: data
      });
    } catch (error) {
      console.log(error);
    }
  }

  public render(): React.ReactElement<IRequestToGraphAndAzureFxProps> {
    const { returnAzureFx } = this.state;

    return (
      <div className={styles.requestToGraphAndAzureFx}>
        <div className={styles.container}>
          <PrimaryButton text="¡PúlsamEE!" onClick={this.clickme} />
          <p>{returnAzureFx}</p>
        </div>
      </div>
    );
  }
}
