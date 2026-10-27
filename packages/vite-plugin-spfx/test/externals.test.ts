import fs from "node:fs";
import path from "node:path";
import { beforeAll, describe, expect, it } from "vitest";
import { createComponentLookup } from "../src/spfx/externals.ts";
import { buildFixture, createFixture, parseDefine, readEntry, readManifest } from "./_fixture.ts";

/**
 * A component must stay external however the import reaches it. Bundling one
 * pulls in SharePoint's own internals — in the reported case sp-http-base and
 * its dynamic import of a Microsoft-internal package, which broke the build.
 */
describe("externalising components at build time", () => {
  let dist: string;

  beforeAll(async () => {
    const root = createFixture({
      components: [
        // npm hoists sp-http next to sp-webpart-base, so a project can import it
        // without declaring it. Generated SPFx projects routinely do.
        {
          name: "@microsoft/sp-http",
          id: "c07208f0-aaaa",
          version: "1.23.2",
          undeclared: true,
          code: `export const SPHttpClient = {};\nexport const broker = () => import("@msinternal/broker");`,
        },
        {
          name: "helper",
          id: "",
          version: "1.0.0",
          plain: true,
          code: `export { PageContext } from "@microsoft/sp-page-context";`,
        },
        {
          name: "@microsoft/sp-page-context",
          id: "1c4541f7-bbbb",
          version: "1.23.2",
          parent: "helper",
          code: `export const PageContext = {};`,
        },
      ],
      entry: [
        `import { SPHttpClient } from "@microsoft/sp-http";`,
        `import { PageContext } from "helper";`,
        `export default class DemoWebPart { client = SPHttpClient; context = PageContext; }`,
      ].join("\n"),
    });
    dist = await buildFixture(root);
  }, 60_000);

  it("keeps a component external that the project only has transitively", () => {
    expect(parseDefine(readEntry(dist).code).deps).toContain("@microsoft/sp-http");
    expect(readManifest(dist).loaderConfig.scriptResources["@microsoft/sp-http"]).toEqual({
      type: "component",
      id: "c07208f0-aaaa",
      version: "1.23.2",
    });
  });

  it("keeps a component external that a bundled package imports from its own node_modules", () => {
    expect(parseDefine(readEntry(dist).code).deps).toContain("@microsoft/sp-page-context");
    expect(readManifest(dist).loaderConfig.scriptResources["@microsoft/sp-page-context"]).toEqual({
      type: "component",
      id: "1c4541f7-bbbb",
      version: "1.23.2",
    });
  });

  it("rejects two versions of a component the project does not pin", async () => {
    // pnpm keeps each dependency's copy apart. With none of its own, the
    // project has no say in which version scriptResources should name.
    const sharedHttp = `export { SPHttpClient } from "@microsoft/sp-http";`;
    const root = createFixture({
      components: [
        { name: "lib-a", id: "", version: "1.0.0", plain: true, code: sharedHttp },
        { name: "lib-b", id: "", version: "1.0.0", plain: true, code: sharedHttp },
        { name: "@microsoft/sp-http", id: "c07208f0-aaaa", version: "1.20.0", parent: "lib-a" },
        { name: "@microsoft/sp-http", id: "c07208f0-aaaa", version: "1.23.2", parent: "lib-b" },
      ],
      entry: [
        `import { SPHttpClient as A } from "lib-a";`,
        `import { SPHttpClient as B } from "lib-b";`,
        `export default class DemoWebPart { clients = [A, B]; }`,
      ].join("\n"),
    });
    await expect(buildFixture(root)).rejects.toThrow(/installed as both/);
  }, 60_000);
});

/**
 * The lookup is what keeps the plugin working across SPFx releases: component
 * ids and versions are read from the installed packages instead of being
 * hardcoded. These tests pin the rules that reading follows.
 */
describe("createComponentLookup", () => {
  it("finds a package by the dist manifest SPFx packages ship", () => {
    const root = createFixture({
      components: [{ name: "@microsoft/sp-core-library", id: "7263c7d0-aaaa", version: "1.23.2" }],
    });
    expect(createComponentLookup(root)("@microsoft/sp-core-library")).toEqual({
      id: "7263c7d0-aaaa",
      version: "1.23.2",
    });
  });

  it("ignores the assembly manifest that sits next to a component", () => {
    // sp-loader ships two manifests; the assembly points at the component via
    // rootComponentId and is not importable. Treating it as a component would
    // put the wrong id into scriptResources.
    const root = createFixture({
      components: [
        {
          name: "@microsoft/sp-loader",
          id: "1c6c9123-bbbb",
          version: "1.23.2",
          withAssembly: true,
        },
      ],
    });
    expect(createComponentLookup(root)("@microsoft/sp-loader")).toEqual({
      id: "1c6c9123-bbbb",
      version: "1.23.2",
    });
  });

  it("skips packages without a dist manifest, so they get bundled", () => {
    const root = createFixture({
      components: [{ name: "lodash", id: "", version: "4", plain: true }],
    });
    expect(createComponentLookup(root)("lodash")).toBeUndefined();
  });

  it("ignores a manifest that is not named after its id", () => {
    // Every imported package is checked, so another tool's app.manifest.json
    // must not turn an ordinary library into an external that SharePoint
    // cannot serve.
    const root = createFixture({
      components: [{ name: "some-lib", id: "", version: "1.0.0", plain: true }],
    });
    const dist = path.join(root, "node_modules/some-lib/dist");
    fs.mkdirSync(dist);
    fs.writeFileSync(path.join(dist, "app.manifest.json"), JSON.stringify({ id: "my-app" }));
    expect(createComponentLookup(root)("some-lib")).toBeUndefined();
  });

  it("uses SPFx's hardcoded ids for the packages that ship no manifest", () => {
    // react, react-dom and microsoft-graph-client are components at runtime but
    // carry no discoverable manifest. SPFx hardcodes exactly these three; if a
    // future release changes an id, this test is where it surfaces.
    const root = createFixture({
      components: [
        { name: "react", id: "", version: "17.0.1", plain: true },
        { name: "react-dom", id: "", version: "17.0.1", plain: true },
      ],
    });
    const findComponent = createComponentLookup(root);
    expect(findComponent("react")).toEqual({
      id: "0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d",
      version: "17.0.1",
    });
    expect(findComponent("react-dom")).toEqual({
      id: "aa0a46ec-1505-43cd-a44a-93f3a5aa460a",
      version: "17.0.1",
    });
  });

  it("strips prerelease suffixes, which SPFx manifests never carry", () => {
    const root = createFixture({
      components: [{ name: "react", id: "", version: "18.3.0-canary.1", plain: true }],
    });
    expect(createComponentLookup(root)("react")?.version).toBe("18.3.0");
  });

  it("prefers the project's copy over one nested in a dependency", () => {
    // scriptResources holds one version per name. Taking whichever copy the
    // first import happens to resolve would make the manifest depend on
    // Rolldown's traversal order.
    const root = createFixture({
      components: [
        { name: "@microsoft/sp-webpart-base", id: "974a7777-aaaa", version: "1.23.2" },
        { name: "@microsoft/sp-lodash-subset", id: "f8a8ad94-bbbb", version: "1.23.0" },
        {
          name: "@microsoft/sp-lodash-subset",
          id: "f8a8ad94-bbbb",
          version: "1.23.2",
          parent: "@microsoft/sp-webpart-base",
        },
      ],
    });
    const importer = path.join(root, "node_modules/@microsoft/sp-webpart-base/index.js");
    expect(createComponentLookup(root)("@microsoft/sp-lodash-subset", importer)?.version).toBe(
      "1.23.0",
    );
  });
});
