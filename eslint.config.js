const spfxProfile = require("@microsoft/eslint-config-spfx/lib/flat-profiles/default");

module.exports = [
  ...spfxProfile,
  {
    files: ["**/*.ts", "**/*.tsx"],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: "./tsconfig.json",
      },
    },
    rules: {
      //implicit is fine.
      "@typescript-eslint/explicit-function-return-type": 0,
    },
  },
];
