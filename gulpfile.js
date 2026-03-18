"use strict";
const fs = require("fs");
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
  //https://github.com/SharePoint/sp-dev-docs/issues/10205
  // generateCssClassName: (name) => {
  //   return name;
  // },
  additionalConfiguration: (generatedConfiguration) => {
    // generatedConfiguration.optimization.usedExports = true;
    //generatedConfiguration.resolve.extensions.push(".ts");
    
    // const sourceMapLoaderIndex = generatedConfiguration.module.rules.findIndex(
    //   (l) => l.use && l.use?.loader?.indexOf("source-map-loader") > -1
    // );
    // if (sourceMapLoaderIndex > -1) {
    //   generatedConfiguration.module.rules.splice(sourceMapLoaderIndex, 1);
    //   generatedConfiguration.module.rules.push({
    //     test: /\.js$/,
    //     enforce: "pre",
    //     use: [
    //       {
    //         loader: "source-map-loader",
    //         options: {
    //           filterSourceMappingUrl: (url, resourcePath) => {
    //             const includeSourceMap =
    //               resourcePath.indexOf("node_modules") === -1 ||
    //               url.indexOf("__spfxCore.js") > -1 ||
    //               url.indexOf("@spfx-extensions/core") > -1 ||
    //               resourcePath.indexOf("@spfx-extensions/core") > -1;
    //             return includeSourceMap;
    //           },
    //         },
    //       },
    //     ],
    //   });
    // }

    
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
