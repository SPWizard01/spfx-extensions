import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderProvider
} from '@microsoft/sp-application-base';
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

  public async onInit(): Promise<void> {
    const envType =
      Environment.type === EnvironmentType.SharePoint
        ? "SharePoint"
        : "ClassicSharePoint";
    const { initCore } = await import(/* webpackChunkName: "spfx-extension-loader" */"../../services/initCoreService");
    //init core then do stuff
    await initCore(
      envType,
      this.context.placeholderProvider,
      this
    );
  }
  protected onPlaceholdersChanged(placeholderProvider: PlaceholderProvider): void {
    window.__SPFxExtensions.Apps.forEach((app) => {
      app.instances.forEach((instance) => {
        instance.executeListeners("onPlaceholdersChanged", placeholderProvider)
      });
    });
  }
  protected onDispose(): void {
    window.__SPFxExtensions.Apps.forEach((app) => {
      app.instances.forEach((instance) => {
        instance.executeListeners("onAppCustomizerDisposed", undefined);
      });
    });
  }
}
