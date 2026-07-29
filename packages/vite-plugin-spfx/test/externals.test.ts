import { describe, expect, it } from "vitest";
import { discoverSpfxExternals } from "../src/spfx/externals.ts";
import { createFixture } from "./_fixture.ts";

/**
 * Discovery is what keeps the plugin working across SPFx releases: component
 * ids and versions are read from the installed packages instead of being
 * hardcoded. These tests pin the rules that reading follows.
 */
describe("discoverSpfxExternals", () => {
  it("finds a package by the dist manifest SPFx packages ship", () => {
    const root = createFixture({
      components: [{ name: "@microsoft/sp-core-library", id: "7263c7d0-aaaa", version: "1.23.2" }],
    });
    expect(discoverSpfxExternals(root).get("@microsoft/sp-core-library")).toEqual({
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
    expect(discoverSpfxExternals(root).get("@microsoft/sp-loader")).toEqual({
      id: "1c6c9123-bbbb",
      version: "1.23.2",
    });
  });

  it("skips packages without a dist manifest, so they get bundled", () => {
    const root = createFixture({
      components: [{ name: "lodash", id: "", version: "4", plain: true }],
    });
    expect(discoverSpfxExternals(root).has("lodash")).toBe(false);
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
    const found = discoverSpfxExternals(root);
    expect(found.get("react")).toEqual({
      id: "0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d",
      version: "17.0.1",
    });
    expect(found.get("react-dom")).toEqual({
      id: "aa0a46ec-1505-43cd-a44a-93f3a5aa460a",
      version: "17.0.1",
    });
  });

  it("strips prerelease suffixes, which SPFx manifests never carry", () => {
    const root = createFixture({
      components: [{ name: "react", id: "", version: "18.3.0-canary.1", plain: true }],
    });
    expect(discoverSpfxExternals(root).get("react")?.version).toBe("18.3.0");
  });
});
