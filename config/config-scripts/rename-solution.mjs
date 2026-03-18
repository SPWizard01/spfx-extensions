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
    console.log("Old Webpart Name:", oldName);
    console.log("New Webpart Name:", newName);
    webpartDataJson.preconfiguredEntries[0].title.default = newName;
    console.log("Updated webpart manifest with new name.");
    console.log(webpartDataJson);
    // await writeFile(jsonPath, JSON.stringify(webpartDataJson, null, 2));
  }
}
