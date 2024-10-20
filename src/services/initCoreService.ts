import type { PageContext } from "@microsoft/sp-page-context";
import {
  CompatibleEnvironmentType,
  SPFxExtensionUtilsPlaceHolderProvider,
} from "spfx-extensions-core";
import { getClassicDisplayMode, getModernDisplayMode } from "spfx-extensions-core/utils/display";
import { loadCoreForSPFxOrClassicWrapper } from "spfx-extensions-core/spfx";
import { PlaceholderProvider } from "@microsoft/sp-application-base";
import { ISPEventObserver } from "@microsoft/sp-core-library";
import { SPFXPREFIX } from "../utilities/constants";
import { getAllConfiguration } from "./idbService";
import { ensureConfigurationList, getConfigurationListData } from "./configurationService";

export async function initCore(
  ctx: PageContext,
  envType: CompatibleEnvironmentType,
  plcHolderProvider?: PlaceholderProvider,
  evtObserver?: ISPEventObserver
) {
  let config = await getAllConfiguration();
  if (!config || config.length === 0) {
    await ensureConfigurationList();
    config = await getConfigurationListData();
  }
  if (!window.__SPFxExtensions) {
    //eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window.__SPFxExtensions as any) = {
      __CoreConfig: config,
    };
  }
  if (!window.__SPFxExtensions.Utils) {
    const buildDate = BUILD_DATE;
    console.info(SPFXPREFIX, "Initializing Core from SPFx Built:", buildDate);

    let resolver: (obj: SPFxExtensionUtilsPlaceHolderProvider) => void;

    const promise = new Promise<SPFxExtensionUtilsPlaceHolderProvider>((resolve) => {
      resolver = resolve;
    });

    let spAppInitializationPromiseResolver = () => {
      // This does nothing. Comment to avoid eslint error
    };

    const spAppInitializationPromise = new Promise<void>((resolve) => {
      spAppInitializationPromiseResolver = resolve;
    });

    const initDispMode =
      envType === "ClassicSharePoint"
        ? getClassicDisplayMode()
        : getModernDisplayMode();
    window.__SPFxExtensions.Utils = {
      //eslint-disable-next-line @typescript-eslint/no-explicit-any
      context: ctx as any,
      environmentType: envType,
      placeHolderProviderPromise: promise,
      placeHolderResolver: resolver!,
      appManifestPromises: [],
      spAppInitializationPromise,
      spAppInitializationPromiseResolver,
      placeHolderResolved: false,
      displayMode: initDispMode,
      initedThroughModern: true,
      fluentIconsInitialized: false,
    };
  }

  //only repopulate if these were not initialized through modern context
  if (window.__SPFxExtensions.Utils && !window.__SPFxExtensions.Utils.initedThroughModern) {
    //eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window.__SPFxExtensions.Utils.context = ctx as any);
    (window.__SPFxExtensions.Utils.environmentType = envType);
    window.__SPFxExtensions.Utils.initedThroughModern = true;
  }
  //resolve once, i.e. modern webpart was loaded before app customizer
  if (
    plcHolderProvider &&
    evtObserver &&
    !window.__SPFxExtensions.Utils.placeHolderResolved
  ) {
    window.__SPFxExtensions.Utils.placeHolderResolved = true;
    window.__SPFxExtensions.Utils.placeHolderResolver({
      //eslint-disable-next-line @typescript-eslint/no-explicit-any
      placeHolderProvider: plcHolderProvider as any,
      //eslint-disable-next-line @typescript-eslint/no-explicit-any
      eventObserver: evtObserver as any,
    });
  }
  return loadCoreForSPFxOrClassicWrapper();
}
