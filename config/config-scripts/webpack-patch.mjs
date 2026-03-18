const isProd =
  process.argv.indexOf("--ship") > -1 ||
  process.argv.indexOf("--production") > -1;
/**
 *
 * @param {import("webpack").Configuration} webpackConfig
 * @returns
 */
//module.exports = function (webpackConfig) {
export default function updatedWebpackConfig(webpackConfig) {
  webpackConfig.output.chunkFilename = "[name]_[contenthash].js";
  webpackConfig.output.environment = {
    arrowFunction: true,
    const: true,
    optionalChaining: true,
    module: true,
    templateLiteral: true,
    destructuring: true,
    dynamicImport: true,
    globalThis: true,
    forOf: true,
    asyncFunction: true,
  };
  // webpackConfig.devtool = "source-map";
  webpackConfig.optimization.splitChunks = {
    cacheGroups: {
      defaultVendors: false,
    },
  };
  // webpackConfig.optimization.minimizer = [];
  //find the source-map-loader and disable
  // const srcMapLoaderIdx = webpackConfig.module.rules.findIndex(
  //   (rule) =>
  //     rule.use?.loader && rule.use.loader.indexOf("source-map-loader") !== -1,
  // );
  // if (srcMapLoaderIdx > -1) {
  //   webpackConfig.module.rules.splice(srcMapLoaderIdx, 1);
  // }
  // console.log(webpackConfig.optimization.minimizer)
  // webpackConfig.optimization.minimizer = [
  //   // new TerserPlugin({
  //   //   exclude:
  //   //     /(spfx-extension-core|spfx-extension-wrapper|spfx-extension-coreconfigurator)/,
  //   // }),
  // ];
  // webpackConfig.module.rules.push({
  //   test: /(spfx-extension-core|spfx-extension-wrapper|spfx-extension-coreconfigurator)/,
  //   // test: /\.js$/,
  //   enforce: "post",
  //   use: { "loader": "source-map-loader" },
  //   // extractSourceMap: true,
  // });
  // webpackConfig.plugins.push(new webpack.SourceMapDevToolPlugin({
  //   test: /spfx-extensions-core\.js/,
  //   // test(asset) {
  //   //   const isMatch = /spfx-extensions-core\.js(\?v=\w+)?$/.test(asset);
  //   //   console.log("Processing asset for source map:", asset, isMatch);
  //   //   return isMatch;
  //   // },
  //   append: `\n//# sourceMappingURL=spfx-extensions-core.js.map`,
  // }))
  // console.log("Updated webpack configuration:", webpackConfig.module.rules);
  webpackConfig.module.rules.push({
    test: /spfx-extensions-core\.js$/,
    generator: {
      filename: "spfx-extensions-core[ext]?v=[hash]",
    },
    type: "asset/resource",
  });
  // webpackConfig.module.rules.push({
  //   test: /\.map$/,
  //   generator: {
  //     filename: "[name][ext]?v=[hash]",
  //   },
  //   type: "asset/resource",
  // });

  webpackConfig.module.rules.push({
    test: /spfx-extensions-coreconfigurator\.js$/,
    generator: {
      filename: "spfx-extensions-coreconfigurator[ext]?v=[hash]",
    },
    type: "asset/resource",
  });
    webpackConfig.module.rules.push({
    test: /spfx-extensions-classiccustomaction\.js$/,
    generator: {
      filename: "spfx-extensions-classiccustomaction[ext]?v=[hash]",
    },
    type: "asset/resource",
  });
  webpackConfig.module.rules.push({
    test: /spfx-extensions-classicwrapper\.js$/,
    generator: {
      filename: "spfx-extensions-classicwrapper[ext]?v=[hash]",
    },
    type: "asset/resource",
  });

  // webpackConfig.resolve.alias["spfx-extensions-core.js"] =
  //   "@spfx-extensions/core/spfx-extensions-core";
  // webpackConfig.resolve.alias["spfx-extensions-core.js.map"] =
  //   "@spfx-extensions/core/dist/core/spfx-extensions-core.js.map";

  // webpackConfig.resolve.alias["spfx-extensions-coreconfigurator.js"] =
  //   "@spfx-extensions/core/spfx-extensions-coreconfigurator";
  // webpackConfig.resolve.alias["spfx-extensions-coreconfigurator.js.map"] =
  //   "@spfx-extensions/core/spfx-extensions-coreconfigurator";

  // webpackConfig.resolve.alias["spfx-extensions-classicwrapper.js"] =
  //   "@spfx-extensions/core/spfx-extensions-classicwrapper";
  // webpackConfig.resolve.alias.push({

  // })
  // webpackConfig.experiments = {
  //   outputModule: true,
  // };
  // webpackConfig.output.libraryTarget = "commonjs-module";
  // webpackConfig.output.library = {
  //   type: "commonjs-module",
  // };
  const definePlugin = webpackConfig.plugins.find((p) => p.definitions);
  if (definePlugin) {
    const date = new Date().toISOString();
    console.log("Adding BUILD_DATE to define plugin:", date);
    definePlugin.definitions["BUILD_DATE"] = JSON.stringify(date);
    console.log("Adding ISDEBUG to define plugin:", isProd ? false : true);
    definePlugin.definitions["ISDEBUG"] = JSON.stringify(isProd ? false : true);
  } else {
    console.error("No define plugin found in webpack configuration.");
  }

  return webpackConfig;
}
