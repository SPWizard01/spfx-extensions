import { SPFXPREFIX } from "../utilities/constants";

/**
 * Points to core location, the holy grail that makes everything working.
 * Setting `localStorage["SPFXEXT"]` to a number i.e. `33343` makes it load
 * from `https://localhost:33343/`.
 *
 * Default URL: ```/sites/appcatalog/CDN/SPFxExtensionAppsCore/core.js```
 */
export async function getRootCoreLocation() {
    const devPort = Number(localStorage.getItem("SPFXEXT"));
    const coreUrls = {
      core: "",
      configuratorUrl: ""
    }
    if (devPort > 0) {
      coreUrls.core = `https://localhost:${devPort}/__spfxCore.js`;
      coreUrls.configuratorUrl = `https://localhost:${devPort}/__spfxCoreConfigurator.js`;
      return coreUrls;
    }
  
    // this part is intercepted by SPFx Webpack and converted later on
    const coreUrl = await import(/* webpackChunkName: "spfx-extension-core-location" */"__spfxCore.js");
    const configuratorUrl = await import(/* webpackChunkName: "spfx-extension-core-location" */"__spfxCoreConfigurator.js");
    if (!coreUrl.default) {
      const msg = "Unable to resolve SPFx Core location";
      throw new Error(`${SPFXPREFIX} ${msg}`);
    }
    if (!configuratorUrl.default) {
      const msg = "Unable to resolve SPFx Core Configurator location";
      throw new Error(`${SPFXPREFIX} ${msg}`);
    }
    coreUrls.core = coreUrl.default;
    coreUrls.configuratorUrl = configuratorUrl.default;
    return coreUrls;
  }
  
  /**
   * Should only be used inside of SPFx or content script in classic pages on SP
   * @returns Singleton promise that resolves once the core is loaded
   */
  export async function loadCoreForSPFxOrClassicWrapper() {
    if (window.__SPFxExtensions.__CorePromise) {
      return window.__SPFxExtensions.__CorePromise;
    }
    const coreUrl = await getRootCoreLocation();
    window.__SPFxExtensions.__ConfiguratorUrl = coreUrl.configuratorUrl;
    window.__SPFxExtensions.__CorePromise = new Promise((resolve) => {
      window.__SPFxExtensions.__CorePromiseResolver = resolve;
      const coreScript = document.createElement("script");
      coreScript.src = coreUrl.core;
      coreScript.type = "module";
      coreScript.addEventListener("error", (err) => {
        console.error(
          SPFXPREFIX,
          "Catastrophic failure, cannot load SPFxExtensions Core from",
          coreUrl,
          err
        );
      });
      document.head.appendChild(coreScript);
    });
  
    return window.__SPFxExtensions.__CorePromise;
  }
  