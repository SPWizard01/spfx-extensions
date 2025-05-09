"use strict";
const fs = require("fs");
const build = require("@microsoft/sp-build-web");
const { spawnSync } = require("child_process");
const path = require("path");
const isProd = process.argv.indexOf("--ship") > -1;
const changeNameArg = "--change-name";
const argvJoined = process.argv.join(" ");
const changeNameArgIndex = argvJoined.indexOf(changeNameArg);
const shouldChangeName = changeNameArgIndex > -1;
const newNameStart = changeNameArgIndex + changeNameArg.length + 1;
const hasMoreArgsAfterName =
  argvJoined.substring(newNameStart).indexOf("--") > -1;
let newName = "";
if (shouldChangeName) {
  if (hasMoreArgsAfterName > -1) {
    newName = argvJoined.substring(
      newNameStart,
      argvJoined.length - hasMoreArgsAfterName - changeNameArg.length - 1
    );
  } else {
    newName = argvJoined.substring(newNameStart);
  }
}

build.addSuppression(
  `Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`
);

const projectSolutionPackage = fs.readFileSync(
  "./config/package-solution.json",
  { encoding: "utf8" }
);
const projectPackage = fs.readFileSync("./package.json", { encoding: "utf8" });
const projectPackageJson = JSON.parse(projectPackage);
const projectSolutionPackageJson = JSON.parse(projectSolutionPackage);
const projectVersion = `${projectPackageJson.version}.0`;
console.log("Project Version:", projectVersion);
projectSolutionPackageJson.solution.version = projectVersion;
projectSolutionPackageJson.solution.features[0].version = projectVersion;

fs.writeFileSync(
  "./config/package-solution.json",
  JSON.stringify(projectSolutionPackageJson, null, 2)
);

if (newName) {
  const jsonPath =
    "./src/webparts/spfxExtensionloader/SpfxExtensionloaderWebPart.manifest.json";
  const webpartData = fs.readFileSync(jsonPath, { encoding: "utf8" });
  const webpartDataJson = JSON.parse(webpartData);
  const oldName = webpartDataJson.preconfiguredEntries[0].title.default;
  console.log("Old Webpart Name:", oldName);
  console.log("New Webpart Name:", newName);
  webpartDataJson.preconfiguredEntries[0].title.default = newName;
  fs.writeFileSync(
    jsonPath,
    JSON.stringify(webpartDataJson, null, 2)
  );
}

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

    generatedConfiguration.module.rules.push({
      test: /__spfxCore\.js$/,
      generator: {
        filename: "spfx-extension-core[ext]?v=[hash]",
      },
      type: "asset/resource",
    });
    generatedConfiguration.module.rules.push({
      test: /__spfxCoreConfigurator\.js$/,
      generator: {
        filename: "spfx-extension-coreconfigurator[ext]?v=[hash]",
      },
      type: "asset/resource",
    });
    generatedConfiguration.module.rules.push({
      test: /__spfxWrapperClassic\.js$/,
      generator: {
        filename: "spfx-extension-wrapper[ext]?v=[hash]",
      },
      type: "asset/resource",
    });
    // generatedConfiguration.module.rules.push({
    //   test: /__spfxCore\.js\.map$/,
    //   generator: {
    //     filename: "spfx-extension-core.js[ext]?[hash]",
    //   },
    //   type: "asset/resource",
    // });
    generatedConfiguration.resolve.alias["__spfxCore.js"] =
      "@spfx-extensions/core/spfxCoreEntry";
    generatedConfiguration.resolve.alias["__spfxCoreConfigurator.js"] =
      "@spfx-extensions/core/configurator";
    generatedConfiguration.resolve.alias["__spfxWrapperClassic.js"] =
      "@spfx-extensions/core/classicWrapper";
    // generatedConfiguration.resolve.alias["__spfxCore.js.map"] =
    //   "@spfx-extensions/core/spfxCoreEntryMap";
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
