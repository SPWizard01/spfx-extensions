"use strict";

const build = require("@microsoft/sp-build-web");
const isProd = process.argv.indexOf("--ship") > -1;
build.addSuppression(
  `Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`
);
build.addSuppression(
  `Warning - [webpack] No webpack config has been provided. Create a webpack.config.js file or call webpack.setConfig({ configPath: null }) in your gulpfile.`
);

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    const definePlugin = generatedConfiguration.plugins.find(
      (p) => p.definitions
    );
    if (definePlugin) {
      const date = new Date().toISOString();
      console.log("Adding BUILD_DATE to define plugin:", date);
      definePlugin.definitions["BUILD_DATE"] = JSON.stringify(date);
      definePlugin.definitions["ISDEBUG"] = JSON.stringify(
        isProd ? false : true
      );
    } else {
      console.error("No define plugin found in webpack configuration.");
    }
  },
});

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set("serve", result.get("serve-deprecated"));

  return result;
};

build.initialize(require("gulp"));
