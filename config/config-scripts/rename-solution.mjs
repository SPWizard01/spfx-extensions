import { readFile, writeFile } from "node:fs/promises";
/**
 *
 * @param {import("@rushstack/heft").IRunScriptOptions} params
 */
export async function runAsync(params) {
  const newName = process.env["SPFX_EXTENSIONS_SOLUTION_NAME"]?.trim() ?? "";
  if (newName) {
    console.log("New webpart name from environment variable:", newName);
    const jsonPath =
      "./src/webparts/spfxExtensionloader/SpfxExtensionloaderWebPart.manifest.json";
    const webpartData = await readFile(jsonPath, { encoding: "utf8" });
    const webpartDataJson = JSON.parse(webpartData);
    const oldName = webpartDataJson.preconfiguredEntries[0].title.default;
    webpartDataJson.preconfiguredEntries[0].title.default = newName;
    await writeFile(jsonPath, JSON.stringify(webpartDataJson, null, 2));
    console.log("Updated webpart manifest with new name.");
  }
}
