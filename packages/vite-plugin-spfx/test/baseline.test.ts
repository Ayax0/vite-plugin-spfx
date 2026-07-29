import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { unzipSync } from "fflate";
import { createSppkg } from "../src/spfx/solution.ts";
import type { ComponentManifest, ScriptResource } from "../src/types.ts";
import { NON_STANDARD_COMPONENTS } from "../src/spfx/externals.ts";

/**
 * Baseline against Microsoft's own toolchain.
 *
 * The fixtures under `fixtures/spfx-<version>/` are frozen output of a stock
 * Heft/webpack SPFx project — see its README for how to refresh them. They are
 * committed rather than read from a sibling project so this package tests
 * standalone; a baseline that skips when a path is missing protects nothing.
 *
 * These tests assert the *contract*, not file equality: key names, resource
 * types, component ids and package layout. Our own build differs in every value
 * (different GUIDs, different hashes) but must match in shape.
 */

const fixtures = path.join(path.dirname(fileURLToPath(import.meta.url)), "fixtures/spfx-1.23.2");

const reference: ComponentManifest = JSON.parse(
  fs.readFileSync(path.join(fixtures, "component.manifest.json"), "utf8"),
);
const referenceParts: string[] = JSON.parse(
  fs.readFileSync(path.join(fixtures, "solution-parts.json"), "utf8"),
);

describe("baseline: component manifest", () => {
  it("has the loaderConfig keys we produce", () => {
    // A new required field in a future SPFx release surfaces here after the
    // fixture is refreshed.
    expect(Object.keys(reference.loaderConfig!).toSorted()).toEqual([
      "entryModuleId",
      "internalModuleBaseUrls",
      "scriptResources",
    ]);
  });

  it("only uses scriptResource types we can emit", () => {
    const used = new Set(
      Object.values(reference.loaderConfig!.scriptResources).map((resource) => resource.type),
    );
    for (const type of used) {
      expect(["path", "component", "localizedPath"]).toContain(type);
    }
  });

  it("names the entry module after the bundle, not the component", () => {
    // entryModuleId is the config.json bundle key and must also be a
    // scriptResources key — the loader looks the entry up by this name.
    const { entryModuleId, scriptResources } = reference.loaderConfig!;
    expect(scriptResources).toHaveProperty(entryModuleId);
    expect(scriptResources[entryModuleId]!.type).toBe("path");
    expect(entryModuleId).not.toBe(reference.id);
  });

  it("agrees with the component ids we hardcode", () => {
    // react, react-dom and microsoft-graph-client ship no discoverable manifest,
    // so their ids are the only ones the plugin cannot look up. This is the test
    // that catches Microsoft moving one of them.
    const components = Object.entries(reference.loaderConfig!.scriptResources).filter(
      (entry): entry is [string, Extract<ScriptResource, { type: "component" }>] =>
        entry[1].type === "component",
    );

    const checked = components.filter(([name]) => name in NON_STANDARD_COMPONENTS);
    expect(checked.length).toBeGreaterThan(0);

    for (const [name, resource] of checked) {
      expect(NON_STANDARD_COMPONENTS[name]).toBe(resource.id);
    }
  });

  it("versions components as plain semver without a prerelease suffix", () => {
    for (const resource of Object.values(reference.loaderConfig!.scriptResources)) {
      if (resource.type !== "component") continue;
      expect(resource.version).toMatch(/^\d+\.\d+\.\d+$/);
    }
  });

  it("drops $schema and keeps the fields that drive the toolbox", () => {
    for (const field of [
      "id",
      "alias",
      "componentType",
      "version",
      "manifestVersion",
      "preconfiguredEntries",
      "supportedHosts",
      "loaderConfig",
    ]) {
      expect(reference).toHaveProperty(field);
    }
    expect(reference).not.toHaveProperty("$schema");
  });
});

describe("baseline: solution package layout", () => {
  it("builds the same OPC parts as Heft", () => {
    const featureId = "ffffffff-1111-2222-3333-444444444444";
    const componentId = "11111111-2222-3333-4444-555555555555";

    const manifest: ComponentManifest = {
      id: componentId,
      alias: "DemoWebPart",
      componentType: "WebPart",
      version: "1.0.0",
      manifestVersion: 2,
      preconfiguredEntries: [{ title: { default: "Demo" } }],
      loaderConfig: {
        internalModuleBaseUrls: [],
        entryModuleId: "demo",
        scriptResources: { demo: { type: "path", path: "demo.js" } },
      },
    };

    const zip = createSppkg({
      solution: {
        name: "demo",
        id: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
        version: "1.0.0.0",
        includeClientSideAssets: true,
        features: [{ id: featureId, title: "Demo", description: "Demo" }],
      },
      manifests: [manifest],
      assets: [{ name: "demo.js", data: new TextEncoder().encode("define()") }],
    });

    const ours = Object.keys(unzipSync(zip))
      .filter((file) => !file.startsWith("ClientSideAssets/"))
      .map((file) => file.replaceAll(featureId, "<feature>").replaceAll(componentId, "<component>"))
      .toSorted();

    expect(ours).toEqual(referenceParts.toSorted());
  });
});
