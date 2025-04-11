import {
  CompatibleEnvironmentType,
  SPFxExtensionUtilsPlaceHolderProvider,
  getClassicDisplayMode,
  getModernDisplayMode
} from "@spfx-extensions/core";
import { PlaceholderProvider } from "@microsoft/sp-application-base";
import { ISPEventObserver } from "@microsoft/sp-core-library";
import { SPFXPREFIX } from "../utilities/constants";
import { loadCoreForSPFxOrClassic } from "@spfx-extensions/core/spfxLoader"

let placeHolderResolved = false;

async function getSuggestedCoreUrl() {
  // this part is intercepted by SPFx Webpack and converted later on
  const coreUrl = await import(/* webpackChunkName: "spfx-extension-core-location" */"__spfxCore.js");
  const configuratorUrl = await import(/* webpackChunkName: "spfx-extension-core-location" */"__spfxCoreConfigurator.js");
  return { coreUrl, configuratorUrl };
}

export async function initCore(
  envType: CompatibleEnvironmentType,
  plcHolderProvider?: PlaceholderProvider,
  evtObserver?: ISPEventObserver
) {
  if (!window.__SPFxExtensions) {
    //eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window.__SPFxExtensions as any) = {
      __CoreConfig: {},
    };
  }
  if (!window.__SPFxExtensions.Utils) {
    const buildDate = BUILD_DATE;
    console.info(SPFXPREFIX, "Initializing Core from SPFx Built:", buildDate);

    const { promise: placeHolderProviderPromise, resolve: placeHolderResolver } = Promise.withResolvers<SPFxExtensionUtilsPlaceHolderProvider>();
    const { promise: spAppInitializationPromise, resolve: spAppInitializationPromiseResolver } = Promise.withResolvers<void>();

    const initDispMode =
      envType === "ClassicSharePoint"
        ? getClassicDisplayMode()
        : getModernDisplayMode();
    window.__SPFxExtensions.Utils = {
      environmentType: envType,
      placeHolderProviderPromise,
      placeHolderResolver,
      appManifestPromises: [],
      spAppInitializationPromise,
      spAppInitializationPromiseResolver,
      displayMode: initDispMode,
      initedThroughModern: true,
      fluentIconsInitialized: false,
      ConfiguratorUrl: "",
    };
  }

  //only repopulate if these were not initialized through modern context
  if (window.__SPFxExtensions.Utils && !window.__SPFxExtensions.Utils.initedThroughModern) {
    (window.__SPFxExtensions.Utils.environmentType = envType);
    window.__SPFxExtensions.Utils.initedThroughModern = true;
  }
  //resolve once, i.e. modern webpart was loaded before app customizer
  if (plcHolderProvider && evtObserver && !placeHolderResolved) {
    placeHolderResolved = true;
    window.__SPFxExtensions.Utils.placeHolderResolver({
      //eslint-disable-next-line @typescript-eslint/no-explicit-any
      placeHolderProvider: plcHolderProvider as any,
      //eslint-disable-next-line @typescript-eslint/no-explicit-any
      eventObserver: evtObserver as any,
    });
  }
  return loadCoreForSPFxOrClassic(getSuggestedCoreUrl)
}
