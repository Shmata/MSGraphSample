import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PnPJsGraphApiWebPartStrings';
import PnPJsGraphApi from './components/PnPJsGraphApi';
import { IPnPJsGraphApiProps } from './components/IPnPJsGraphApiProps';
import { MSGraphClient } from '@microsoft/sp-http';


export interface IPnPJsGraphApiWebPartProps {
  description: string;
}

export default class PnPJsGraphApiWebPart extends BaseClientSideWebPart<IPnPJsGraphApiWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void =>{
        // get information about the current user from the Microsoft Graph
        client
        //.api('/me/messages')
        .api('/users')
        .get((error, users: any, rawResponse?: any)=>{
          console.log(error, users, rawResponse);  
          this.domElement.innerHTML = `
            <div class="">
              <div>
                  <h3>Welcome to SharePoint Framework!</h3>
                  <p>
                      The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
                  </p>
              </div>
              <div id="spListContainer" />
            </div>`;

            // List the latest emails based on what we got from the Graph
            this._renderEmailList(users.value);
        });
    });



    const element: React.ReactElement<IPnPJsGraphApiProps> = React.createElement(
      PnPJsGraphApi,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _renderEmailList(clinets: any[]): void {
    console.log(clinets);
    let html: string = '';
    // for (let index = 0; index < clinets.length; index++) {
    //   html += `<p class="welcome">Email ${index + 1} - ${escape(clinets[index].mail)}</p>`;
    // }
    clinets.map((val,key)=>{
      html += `<p id='${key}' class="welcome">Email: ${val.mail} </p> <br />`;
    })
    // Add the emails to the placeholder
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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