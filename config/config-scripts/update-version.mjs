import { readFile, writeFile } from "node:fs/promises";
/**
 *
 * @param {import("@rushstack/heft").IRunScriptOptions} params
 */
export async function runAsync(params) {
  const projectSolutionPackage = await readFile(
    "./config/package-solution.json",
    { encoding: "utf8" },
  );
  const projectPackage = await readFile("./package.json", { encoding: "utf8" });
  const projectPackageJson = JSON.parse(projectPackage);
  const projectSolutionPackageJson = JSON.parse(projectSolutionPackage);
  const projectVersion = `${projectPackageJson.version}.0`;
  console.log("Project Version:", projectVersion);
  projectSolutionPackageJson.solution.version = projectVersion;
  projectSolutionPackageJson.solution.features[0].version = projectVersion;

  await writeFile(
    "./config/package-solution.json",
    JSON.stringify(projectSolutionPackageJson, null, 2),
  );
}
