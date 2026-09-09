import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'XsProjectWebPartStrings';
import XsProject from './components/XsProject/XsProject';
import { IXsProjectProps } from './components/XsProject/IXsProjectProps';
import { initializeEnvironment } from '../../config/environment';
import { IApiServiceConfiguration } from '../../models/ApiResponse';
import { IEnvironmentConfiguration } from '../../models/Environment';
import { IServiceContainer, ServiceFactory } from '../../services/ServiceFactory';

const LOG_SOURCE: string = 'XsProjectWebPart';

export interface IXsProjectWebPartProps {
  description: string;
  /** Optional per-instance override of `api.baseUrl` from `src/config/environment.ts`. */
  apiBaseUrl: string;
  /** Optional per-instance override of `api.resourceUri` from `src/config/environment.ts`. */
  apiResourceUri: string;
}

export default class XsProjectWebPart extends BaseClientSideWebPart<IXsProjectWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const services: IServiceContainer = this._getServices();

    const element: React.ReactElement<IXsProjectProps> = React.createElement(
      XsProject,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        userDisplayName: this.context.pageContext.user.displayName,
        projectService: services.projectService,
        buildingService: services.buildingService,
        lookupService: services.lookupService,
        loginService: services.loginService,
        sharePointSiteService: services.sharePointSiteService
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /**
   * Resolves the service layer.
   *
   * Settings come from `src/config/environment.ts`; the property pane fields are only
   * per-instance overrides and stay empty in normal use. This is the only place in the
   * web part that knows about concrete service classes - components receive the
   * resolved interfaces as props.
   */
  private _getServices(): IServiceContainer {
    const overrides: IApiServiceConfiguration = {
      baseUrl: (this.properties.apiBaseUrl || '').trim() || undefined,
      resourceUri: (this.properties.apiResourceUri || '').trim() || undefined
    };

    return ServiceFactory.getServices(this.context, overrides);
  }

  protected onInit(): Promise<void> {
    // Resolve the environment once, using the SPFx context for exact detection.
    const environment: IEnvironmentConfiguration = initializeEnvironment({
      hostName: this.context.pageContext.web.absoluteUrl,
      isServedFromLocalhost: this.context.isServedFromLocalhost
    });
    Log.info(LOG_SOURCE, `Running against the '${environment.name}' environment (${environment.api.baseUrl}).`);

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
            },
            {
              groupName: strings.ApiGroupName,
              groupFields: [
                PropertyPaneTextField('apiBaseUrl', {
                  label: strings.ApiBaseUrlFieldLabel,
                  description: strings.ApiBaseUrlFieldDescription
                }),
                PropertyPaneTextField('apiResourceUri', {
                  label: strings.ApiResourceUriFieldLabel,
                  description: strings.ApiResourceUriFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
