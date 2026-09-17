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

/**
 * Floor for the app height, used when the measurement below cannot produce a sensible
 * one - a web part rendered while hidden, or before the hosting page has laid out,
 * measures as having no room at all.
 */
const MIN_APP_HEIGHT_PX: number = 480;

/** A canvas ancestor whose fixed width cap was lifted, and the inline value to put back. */
interface IReleasedWidthCap {
  element: HTMLElement;
  maxWidth: string;
}

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
  /** Canvas ancestors this web part widened, so `onDispose` can put the page back. */
  private _releasedWidthCaps: IReleasedWidthCap[] = [];
  private _isTrackingResize: boolean = false;

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
    this._fillAvailableSpace();
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

  /**
   * Lets the app use the whole content area, in both directions.
   *
   * Nothing inside the bundle can do this on its own. `domElement` is nested in
   * SharePoint's canvas (`.CanvasZone` / `.CanvasSection` / `.ControlZone`), and a normal
   * content section caps that chain at a fixed pixel width and centres it with auto
   * margins - which is what made this app a narrow column in the middle of the page, no
   * matter what its own stylesheets said. `supportsFullBleed` in the manifest lets an
   * author put the web part in a full-width section instead, but it still has to fill a
   * normal one, so the cap is lifted here as well.
   */
  private _fillAvailableSpace(): void {
    this._releaseWidthCaps();
    this._measureAvailableHeight();

    // Registered once, however often SharePoint re-renders the web part.
    if (!this._isTrackingResize) {
      window.addEventListener('resize', this._measureAvailableHeight);
      this._isTrackingResize = true;
    }
  }

  /**
   * Lifts the fixed width cap off this web part's own canvas ancestors.
   *
   * Deliberately narrow in what it touches: only ancestors of *this* web part, only their
   * `max-width`, and only when that resolves to a fixed pixel value - a percentage cap is
   * a deliberate proportion of a parent that is itself being widened here, not a centring
   * device, so it is left alone. Every original inline value is remembered and put back in
   * `onDispose`, so removing the web part leaves the page exactly as it was, and no other
   * section or web part on the page is affected either way.
   */
  private _releaseWidthCaps(): void {
    // A re-render re-reads the page rather than stacking overrides on top of its own.
    this._restoreWidthCaps();

    let ancestor: HTMLElement | null = this.domElement.parentElement;

    while (ancestor && ancestor !== document.body) {
      const maxWidth: string = window.getComputedStyle(ancestor).maxWidth;

      if (/px$/.test(maxWidth)) {
        this._releasedWidthCaps.push({ element: ancestor, maxWidth: ancestor.style.maxWidth });
        ancestor.style.maxWidth = 'none';
      }

      ancestor = ancestor.parentElement;
    }
  }

  private _restoreWidthCaps(): void {
    this._releasedWidthCaps.forEach((cap: IReleasedWidthCap) => {
      cap.element.style.maxWidth = cap.maxWidth;
    });

    this._releasedWidthCaps = [];
  }

  /**
   * Publishes the height left below the web part as `--xsProjectAvailableHeight`, which
   * `XsProject.module.scss` gives the app shell as a `min-height`.
   *
   * It has to be measured rather than written as `100vh` in the stylesheet: the web part
   * starts below SharePoint's own page chrome, whose height differs between a team site, a
   * communication site, a Teams tab and the workbench - and changes again when the page is
   * put into edit mode. Assigned as a property so the `this` pointer survives being handed
   * to `addEventListener`.
   */
  private _measureAvailableHeight = (): void => {
    const offsetTop: number = this.domElement.getBoundingClientRect().top + (window.scrollY || 0);
    const available: number = Math.max(MIN_APP_HEIGHT_PX, window.innerHeight - offsetTop);

    this.domElement.style.setProperty('--xsProjectAvailableHeight', `${Math.round(available)}px`);
  };

  protected onDispose(): void {
    window.removeEventListener('resize', this._measureAvailableHeight);
    this._isTrackingResize = false;
    // The page outlives this web part, so it gets its own width back.
    this._restoreWidthCaps();
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
