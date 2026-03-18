import { CompatibleEnvironmentType } from "@spfx-extensions/core";
import { PlaceholderProvider } from "@microsoft/sp-application-base";
import { ISPEventObserver } from "@microsoft/sp-core-library";
import { SPFXPREFIX } from "../utilities/constants";
import { loadCoreForSPFxOrClassic } from "@spfx-extensions/core/spfx-extensions-loader";
import coreUrl from "@spfx-extensions/core/spfx-extensions-core";
import configuratorUrl from "@spfx-extensions/core/spfx-extensions-coreconfigurator";
import wrapperUrl from "@spfx-extensions/core/spfx-extensions-classicwrapper";
import customActionUrl from "@spfx-extensions/core/spfx-extensions-classiccustomaction";

let placeHolderResolved = false;

async function getSuggestedCoreUrl() {
  // return Promise.resolve({ coreUrl, configuratorUrl, wrapperUrl });
  return Promise.resolve({
    coreUrl: coreUrl as string,
    configuratorUrl: configuratorUrl as unknown as string,
    wrapperUrl: wrapperUrl as unknown as string,
    customActionUrl: customActionUrl as unknown as string,
  });
}

/**
 * WILL FAIL IF CALLED BEFORE `initCore`
 *
 * Registers a placeholder provider to be used by the core library
 * @returns void
 */
export function registerPlaceHolderProvider(
  placeHolderProvider: PlaceholderProvider,
  eventObserver: ISPEventObserver,
) {
  //resolve once, i.e. modern webpart was loaded before app customizer
  if (!placeHolderResolved) {
    placeHolderResolved = true;
    window.__SPFxExtensions.Utils.placeHolderResolver({
      //eslint-disable-next-line @typescript-eslint/no-explicit-any
      placeHolderProvider: placeHolderProvider as any,
      //eslint-disable-next-line @typescript-eslint/no-explicit-any
      eventObserver: eventObserver as any,
    });
  }
}

export async function initCore(envType: CompatibleEnvironmentType) {
  if (window.__SPFxExtensions?.__CoreInitializationPromise) {
    //only repopulate if these were not initialized through modern context
    if (
      window.__SPFxExtensions.Utils &&
      !window.__SPFxExtensions.Utils.initializedThroughSPFX
    ) {
      window.__SPFxExtensions.Utils.environmentType = envType;
      window.__SPFxExtensions.Utils.initializedThroughSPFX = true;
    }
    return window.__SPFxExtensions.__CoreInitializationPromise;
  }

  const { promise, resolve } = Promise.withResolvers<void>();
  //eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window.__SPFxExtensions as any) = {
    __CoreInitializationPromise: promise,
    __CoreInitializationResolver: resolve,
  };
  const buildDate = BUILD_DATE;
  console.info(SPFXPREFIX, "Initializing Core from SPFx Built:", buildDate);
  await loadCoreForSPFxOrClassic(getSuggestedCoreUrl, envType, true);
  window.__SPFxExtensions.__CoreInitializationResolver();
}
