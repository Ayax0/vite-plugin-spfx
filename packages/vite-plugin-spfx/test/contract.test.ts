import fs from "node:fs";
import path from "node:path";
import { beforeAll, describe, expect, it } from "vitest";
import { init, parse } from "es-module-lexer";
import { unzipSync } from "fflate";
import {
  buildFixture,
  createFixture,
  DEFAULT_MANIFEST,
  parseDefine,
  readDist,
  readEntry,
  readManifest,
} from "./_fixture.ts";

/**
 * The contract the SharePoint loader imposes, asserted against a real
 * production build.
 *
 * Every bug this project hit passed in dev and broke in production, because
 * minification runs between the hooks. Unit tests alone would have missed all
 * of them, so this suite builds the thing that actually ships.
 */
describe("SharePoint loader contract", () => {
  let dist: string;

  beforeAll(async () => {
    const root = createFixture({
      components: [
        { name: "@microsoft/sp-webpart-base", id: "974a7777-aaaa", version: "1.23.2" },
        { name: "@microsoft/sp-core-library", id: "7263c7d0-bbbb", version: "1.23.2" },
      ],
      entry: [
        `import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";`,
        `import { Version } from "@microsoft/sp-core-library";`,
        `import * as strings from "DemoStrings";`,
        `import "./styles.css";`,
        `export default class DemoWebPart extends BaseClientSideWebPart {`,
        `  render() { this.domElement.innerHTML = strings.Greeting + Version; }`,
        `}`,
      ].join("\n"),
      styles: `.demo { color: "[theme:bodyText, default: #323130]"; color: var(--bodyText); }`,
      locales: { "en-us": `define([], function () { return { Greeting: "hi" }; });` },
    });
    dist = await buildFixture(root, {
      localizedResources: { DemoStrings: "src/loc/{locale}.js" },
    });
  }, 60_000);

  it("emits one named define whose id is <componentId>_<version>", () => {
    // RequireJS resolves the module by this exact name; a mismatch makes the
    // loader report "loaded but did not return" rather than a clear error.
    // The version comes from package.json because the manifest says "*".
    const { code } = readEntry(dist);
    expect(parseDefine(code).id).toBe(`${DEFAULT_MANIFEST.id}_1.2.3`);
    expect(code.match(/\bdefine\(/g)).toHaveLength(1);
  });

  it("ships a classic script with no ESM syntax left", async () => {
    await init;
    const { code } = readEntry(dist);
    const [imports, exports] = parse(code);
    expect(imports.filter((i) => i.n !== undefined)).toHaveLength(0);
    expect(exports).toHaveLength(0);
  });

  it("lists exactly the scriptResources keys as dependencies, in order", () => {
    // The loader maps the define array onto scriptResources positionally, so an
    // extra, missing or reordered entry silently binds the wrong module.
    const { deps: declared } = parseDefine(readEntry(dist).code);
    const { loaderConfig } = readManifest(dist);
    const resources = Object.keys(loaderConfig.scriptResources);

    expect(resources[0]).toBe(loaderConfig.entryModuleId);
    expect(resources.slice(1)).toEqual(declared);
  });

  it("points the entry resource at the file that was actually emitted", () => {
    const { fileName } = readEntry(dist);
    const { loaderConfig } = readManifest(dist);
    expect(loaderConfig.scriptResources[loaderConfig.entryModuleId]).toEqual({
      type: "path",
      path: fileName,
    });
    // The loader strips ".js" and lets RequireJS re-append it.
    expect(fileName.endsWith(".js")).toBe(true);
  });

  it("resolves SPFx packages to components and strings to a path", () => {
    const { scriptResources } = readManifest(dist).loaderConfig;
    expect(scriptResources["@microsoft/sp-webpart-base"]).toEqual({
      type: "component",
      id: "974a7777-aaaa",
      version: "1.23.2",
    });
    expect(scriptResources["DemoStrings"].type).toBe("path");
    expect(readDist(dist)).toContain(scriptResources["DemoStrings"].path);
  });

  it("ships no stylesheet, because the manifest cannot reference one", () => {
    // SPFx has no CSS resource type: styles must be injected from JS or they
    // never load at all.
    expect(readDist(dist).filter((file) => file.endsWith(".css"))).toEqual([]);
    const { code } = readEntry(dist);
    expect(code).toMatch(/createElement\(\s*["'`]style["'`]\s*\)/);
    expect(code).toContain("data-spfx-vite");
  });

  it("keeps theme fallbacks through minification", () => {
    // The rewrite has to run in PostCSS: the minifier removes the themed
    // declaration as redundant, so a later pass would find nothing.
    const { code } = readEntry(dist);
    expect(code).toContain("var(--bodyText,#323130)");
  });

  it("resolves asset URLs against the script's own location", () => {
    const { code } = readEntry(dist);
    expect(code).toContain("document.currentScript");
  });
});

describe("solution package", () => {
  it("contains exactly the assets the manifest references", async () => {
    const root = createFixture({
      components: [{ name: "@microsoft/sp-webpart-base", id: "974a7777-aaaa", version: "1.23.2" }],
      entry: [
        `import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";`,
        `export default class DemoWebPart extends BaseClientSideWebPart {}`,
      ].join("\n"),
      packageSolution: {
        solution: {
          name: "demo-client-side-solution",
          id: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
          version: "1.0.0.0",
          includeClientSideAssets: true,
          skipFeatureDeployment: true,
        },
        paths: { zippedPackage: "solution/demo.sppkg" },
      },
    });
    const dist = await buildFixture(root);

    const sppkg = path.join(root, "sharepoint", "solution", "demo.sppkg");
    expect(fs.existsSync(sppkg)).toBe(true);

    const files = Object.keys(unzipSync(fs.readFileSync(sppkg)));
    const packaged = files
      .filter((name) => name.startsWith("ClientSideAssets/"))
      .map((name) => name.slice("ClientSideAssets/".length));

    // Every referenced path must exist in the package, and nothing else may:
    // a stray .css or .map would be served but never used.
    const { scriptResources } = readManifest(dist).loaderConfig;
    const referenced = Object.values<any>(scriptResources)
      .filter((resource) => resource.type === "path")
      .map((resource) => resource.path);

    expect(packaged.toSorted()).toEqual(referenced.toSorted());
  }, 60_000);
});
