import { loadCoreForSPFxOrClassic } from "@spfx-extensions/core/spfxLoader"

async function getSuggestedCoreUrl() {
  // this part is intercepted by SPFx Webpack and converted later on

  const coreUrl = await import(/* webpackChunkName: "spfx-extension-core-location" */"__spfxCore.js");
  const configuratorUrl = await import(/* webpackChunkName: "spfx-extension-core-location" */"__spfxCoreConfigurator.js");
  return { coreUrl, configuratorUrl };
}

export async function loadCore() {
  return loadCoreForSPFxOrClassic(getSuggestedCoreUrl)
}
