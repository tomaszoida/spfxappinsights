import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { ApplicationInsights, IPageViewTelemetry } from '@microsoft/applicationinsights-web';
import * as strings from 'SpfxappinsightsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'SpfxappinsightsApplicationCustomizer';
const KeyDefaultValue: string = 'AppInsightsInstrumentationKey';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpfxappinsightsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  instrumentationKey: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxappinsightsApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxappinsightsApplicationCustomizerProperties> {

  private _instrumentationKey: string;
  private _appInsights: ApplicationInsights;
  private _startingPage: string;
  @override
  public onInit(): Promise<void> {
    this._startingPage = window.location.href;
    this._instrumentationKey = this.properties.instrumentationKey;

    if (this._instrumentationKey && this._instrumentationKey != KeyDefaultValue) {
      this._appInsights = new ApplicationInsights({
        config: {
          instrumentationKey: this._instrumentationKey
        }
      });

      this._appInsights.loadAppInsights();
      this.context.application.navigatedEvent.add(this, this._navigationEventHandler);
    }

    return Promise.resolve();
  }

  private _navigationEventHandler(): void {
    let isPartialReload = this._startingPage != window.location.href;
    this._trackIntitialPageViewWithContext(this.context.pageContext.legacyPageContext, isPartialReload);
  }

  private _trackIntitialPageViewWithContext(context: any, isPartialReload: boolean) {

    let properties: IPageViewTelemetry = {
      properties: {        
        "sppagecontextinfo.currentCultureName": context.currentCultureName,
        "sppagecontextinfo.currentUICultureName": context.currentUICultureName,
        "sppagecontextinfo.isExternalGuestUser": context.isExternalGuestUser,
        "sppagecontextinfo.isEmailAuthenticationGuestUser": context.isEmailAuthenticationGuestUser,
        "sppagecontextinfo.isAnonymousGuestUser": context.isAnonymousGuestUser,
        "sppagecontextinfo.isSiteOwner": context.isSiteOwner,
        "sppagecontextinfo.isSiteAdmin": context.isSiteAdmin,
        "sppagecontextinfo.webTemplateConfiguration": context.webTemplateConfiguration,
        "sppagecontextinfo.webTitle": context.webTitle,
        "sppagecontextinfo.webAbsoluteUrl": context.webAbsoluteUrl,
        "sppagecontextinfo.siteAbsoluteUrl": context.siteAbsoluteUrl,
        "sppagecontextinfo.listTitle": context.listTitle,
        "snppagecontextinfo.isWebWelcomePage": context.isWebWelcomePage,
        "sharepoint.isPartialReload": isPartialReload,
      }
    };

    if (isPartialReload) {
      // TODO track partial reload performance
      properties.properties.duration = 0;
    }

    this._appInsights.trackPageView(properties);
  }
}
