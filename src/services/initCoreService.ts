import {
  CompatibleEnvironmentType,
} from "@spfx-extensions/core";
import { PlaceholderProvider } from "@microsoft/sp-application-base";
import { ISPEventObserver } from "@microsoft/sp-core-library";
import { SPFXPREFIX } from "../utilities/constants";
import { loadCoreForSPFxOrClassic } from "@spfx-extensions/core/spfxLoader"
import coreUrl from "__spfxCore.js";
import configuratorUrl from "__spfxCoreConfigurator.js";
import wrapperUrl from "__spfxWrapperClassic.js";

let placeHolderResolved = false;
let corePromise: Promise<void> | undefined = undefined;

async function getSuggestedCoreUrl() {
  return Promise.resolve({ coreUrl, configuratorUrl, wrapperUrl });
}

export async function initCore(
  envType: CompatibleEnvironmentType,
  placeHolderProvider?: PlaceholderProvider,
  eventObserver?: ISPEventObserver
) {
  if (!corePromise) {
    const buildDate = BUILD_DATE;
    console.info(SPFXPREFIX, "Initializing Core from SPFx Built:", buildDate);
    corePromise = loadCoreForSPFxOrClassic(getSuggestedCoreUrl, envType, true);
  }
  await corePromise;

  //only repopulate if these were not initialized through modern context
  if (window.__SPFxExtensions.Utils && !window.__SPFxExtensions.Utils.initializedThroughSPFX) {
    window.__SPFxExtensions.Utils.environmentType = envType;
    window.__SPFxExtensions.Utils.initializedThroughSPFX = true;
  }
  //resolve once, i.e. modern webpart was loaded before app customizer
  if (!placeHolderResolved && placeHolderProvider && eventObserver) {
    placeHolderResolved = true;
    window.__SPFxExtensions.Utils.placeHolderResolver({
      //eslint-disable-next-line @typescript-eslint/no-explicit-any
      placeHolderProvider: placeHolderProvider as any,
      //eslint-disable-next-line @typescript-eslint/no-explicit-any
      eventObserver: eventObserver as any,
    });
  }
}
