import SourceMapDevToolPlugin from "webpack/lib/SourceMapDevToolPlugin.js";
import TerserPlugin from "terser-webpack-plugin";
import CopyPlugin from "copy-webpack-plugin";
import path from "path";
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
  webpackConfig.devtool = "source-map";
  webpackConfig.optimization.splitChunks = {
    cacheGroups: {
      defaultVendors: false,
    },
  };
  webpackConfig.optimization.minimizer = [];
  webpackConfig.optimization.minimizer = [
    new TerserPlugin({
      // minify: TerserPlugin.esbuildMinify,
      extractComments: false,
      terserOptions: {
        // comments: "all",
        sourceMap: false,
      },
      exclude:
        /(spfx-extension-core|spfx-extension-wrapper|spfx-extension-coreconfigurator)/,
    }),
  ];
  webpackConfig.module.rules.push({
    test: /(spfx-extension-core|spfx-extension-wrapper|spfx-extension-coreconfigurator)/,
    extractSourceMap: true,
  });
  // console.log(webpackConfig.plugins);
  // webpackConfig.plugins.push(
  //   new CopyPlugin({
  //     patterns: [
  //       {
  //         from: "*.js.map",
  //         context: "node_modules/@spfx-extensions/core/dist/core",
  //         // context: path.relative(path.dirname()),
  //         info: (file) => {
  //           console.log(file);
  //           return { minimized: true };
  //         },
  //       },
  //     ],
  //   }),
  // );
  // webpackConfig.plugins.push(new SourceMapDevToolPlugin({
  //   // test: /__spfxCore\.js$/
  //   test: (asset) =>{
  //     console.log("Evaluating asset for source map generation:", asset);
  //     return false;
  //   }
  // }))
  webpackConfig.module.rules.push({
    test: /__spfxCore\.js$/,
    generator: {
      filename: "spfx-extension-core[ext]?v=[hash]",
    },
    type: "asset/resource",
  });
  webpackConfig.module.rules.push({
    test: /__spfxCoreConfigurator\.js$/,
    generator: {
      filename: "spfx-extension-coreconfigurator[ext]?v=[hash]",
    },
    type: "asset/resource",
  });
  webpackConfig.module.rules.push({
    test: /__spfxWrapperClassic\.js$/,
    generator: {
      filename: "spfx-extension-wrapper[ext]?v=[hash]",
    },
    type: "asset/resource",
  });
  webpackConfig.module.rules.push({
    test: /__spfxCore\.js\.map$/,
    generator: {
      filename: "spfx-extension-core.js[ext]?[hash]",
    },
    type: "asset/resource",
  });
  webpackConfig.resolve.alias["__spfxCore.js"] =
    "@spfx-extensions/core/spfxCoreEntry";
  webpackConfig.resolve.alias["__spfxCoreConfigurator.js"] =
    "@spfx-extensions/core/configurator";
  webpackConfig.resolve.alias["__spfxWrapperClassic.js"] =
    "@spfx-extensions/core/classicWrapper";
  webpackConfig.resolve.alias["__spfxCore.js.map"] =
    "@spfx-extensions/core/dist/core/__spfxCore.js.map";
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
