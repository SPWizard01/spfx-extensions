import {
  BaseApplicationCustomizer,
  PlaceholderProvider
} from '@microsoft/sp-application-base';
import { SPFXPREFIX } from '../../utilities/constants';
//import * as strings from 'SpfxExtensionApplicationCustomizerStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpfxExtensionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  _noProps: boolean;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxExtensionApplicationCustomizerProperties> {

  corePromise = new Promise((resolve) => {
    import(/* webpackChunkName: "spfx-extension-loader" */"../../services/initCoreService").then(({ initCore, registerPlaceHolderProvider }) => {
      initCore("SharePoint").then(() => {
        registerPlaceHolderProvider(this.context.placeholderProvider, this);
        resolve(true);
      }).catch((e) => {
        console.error(SPFXPREFIX, "Error while initializing core from application customizer", e);
      })
    }).catch((e) => {
      console.error(SPFXPREFIX, "Error while importing core from application customizer", e);
    })
  });


  // private async ensureCoreAndRegister() {
  //   const { initCore, registerPlaceHolderProvider } = await import(/* webpackChunkName: "spfx-extension-loader" */"../../services/initCoreService");
  //   await initCore("SharePoint");
  //   registerPlaceHolderProvider(this.context.placeholderProvider, this);
  // }

  public async onInit(): Promise<void> {
    await this.corePromise;
  }

  waitAndNotifyPlaceholderChanged(placeholderProvider: PlaceholderProvider) {
    this.corePromise!.then(() => {
      window.__SPFxExtensions.Apps.forEach((app) => {
        app.instances.forEach((instance) => {
          instance.executeListeners("onPlaceholdersChanged", placeholderProvider)
        });
      });
    }).catch(() => {
      //swallow
    });
  }

  protected onPlaceholdersChanged(placeholderProvider: PlaceholderProvider) {
    // this fires before onInit
    this.waitAndNotifyPlaceholderChanged(placeholderProvider);
  }

  protected onDispose(): void {
    window.__SPFxExtensions?.Apps?.forEach((app) => {
      app.instances.forEach((instance) => {
        instance.executeListeners("onAppCustomizerDisposed", undefined);
      });
    });
  }
}
