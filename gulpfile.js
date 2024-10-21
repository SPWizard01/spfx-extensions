"use strict";

const build = require("@microsoft/sp-build-web");
const { spawnSync } = require("child_process");
const path = require("path");
const isProd = process.argv.indexOf("--ship") > -1;
build.addSuppression(
  `Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`
);
// build.addSuppression(
//   `Warning - [webpack] No webpack config has been provided. Create a webpack.config.js file or call webpack.setConfig({ configPath: null }) in your gulpfile.`
// );
// build.webpack.setConfig({
//   configPath: "./webpack.config.js"
// })

build.tscCmd.executeTask = function () {
  console.log("Overriding tscCmd.executeTask");
  spawnSync("npx", ["tsc", "--build", "--verbose"], {
    shell: true,
    stdio: "inherit",
    cwd: path.resolve(__dirname),
  });
  return Promise.resolve();
};

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    // generatedConfiguration.optimization.usedExports = true;
    //generatedConfiguration.resolve.extensions.push(".ts");
    generatedConfiguration.output.chunkFilename = "[name]_[contenthash].js";
    generatedConfiguration.output.environment = {
      arrowFunction: true,
      const: true,
      optionalChaining: true,
      module: true,
      templateLiteral: true,
      destructuring: true,
      dynamicImport: true,
      globalThis: true,
      forOf: true,
    };
    generatedConfiguration.optimization.splitChunks = {
      cacheGroups: {
        defaultVendors: false,
      },
    };
    generatedConfiguration.module.rules.push({
      test: /__spfxCore.js/,
      generator: {
        filename: "spfx-extension-core_[hash][ext]",
      },
      type: "asset/resource",
    });
    generatedConfiguration.resolve.alias["__spfxCore.js"] =
      "@spfx-extensions/core/core";

    // generatedConfiguration.resolve.alias.push({

    // })
    // generatedConfiguration.experiments = {
    //   outputModule: true,
    // };
    // generatedConfiguration.output.libraryTarget = "commonjs-module";
    // generatedConfiguration.output.library = {
    //   type: "commonjs-module",
    // };
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
    return generatedConfiguration;
  },
});

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set("serve", result.get("serve-deprecated"));

  return result;
};

build.initialize(require("gulp"));
