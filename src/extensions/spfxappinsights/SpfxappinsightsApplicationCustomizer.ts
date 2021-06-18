import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { SPPageContextInfo } from 'sppagecontextinfo';
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

  @override
  public onInit(): Promise<void> {
    let instrumentationKey: string = this.properties.instrumentationKey;
    
    if (instrumentationKey && instrumentationKey != KeyDefaultValue) {
      return SPPageContextInfo.getContext().then(context => {
        const appInsights = new ApplicationInsights({
          config: {
            instrumentationKey: instrumentationKey
          }
        });

        appInsights.loadAppInsights();
        appInsights.trackPageView({
          properties: {
            "document.referrer": document.referrer,
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
          }
        });
      });
    }

    return Promise.resolve();
  }
}
