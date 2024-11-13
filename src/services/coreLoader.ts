/**
 * Points to core location, the holy grail that makes everything working.
 * Setting `localStorage["SPFXEXT"]` to a number i.e. `33343` makes it load
 * from `https://localhost:33343/`.
 *
 * Default URL: ```/sites/appcatalog/CDN/SPFxExtensionAppsCore/core.js```
 */
export async function getRootCoreLocation() {
    const devPort = Number(localStorage.getItem("SPFXEXT"));
    if (devPort > 0) {
      return `https://localhost:${devPort}/__spfxCore.js`;
    }
  
    // this part is intercepted by SPFx Webpack and converted later on
    const webpackCoreUrl = await import(/* webpackChunkName: "spfx-extension-core-location" */"__spfxCore.js");
    //side effect for core map to be included in the build
    //import(/* webpackChunkName: "spfx-extension-core-location" */"__spfxCore.js.map").catch(() => {});
    if (!webpackCoreUrl.default) {
      const msg = "Unable to resolve SPFx Core location";
      console.error(msg, webpackCoreUrl);
      throw new Error(msg);
    }
    return webpackCoreUrl.default;
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
    window.__SPFxExtensions.__CorePromise = new Promise((resolve) => {
      window.__SPFxExtensions.__CorePromiseResolver = resolve;
      const coreScript = document.createElement("script");
      coreScript.src = coreUrl;
      coreScript.type = "module";
      coreScript.addEventListener("error", (err) => {
        console.error(
          "Catastrophic failure, cannot load SPFxExtensions Core from",
          coreUrl,
          err
        );
      });
      document.head.appendChild(coreScript);
    });
  
    return window.__SPFxExtensions.__CorePromise;
  }
  