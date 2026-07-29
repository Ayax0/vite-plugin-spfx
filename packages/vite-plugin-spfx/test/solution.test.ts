import { describe, expect, it } from "vitest";
import { unzipSync, strFromU8 } from "fflate";
import { createSppkg } from "../src/spfx/solution.ts";
import type { ComponentManifest, Solution } from "../src/types.ts";

const solution: Solution = {
  name: "demo-client-side-solution",
  id: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
  version: "1.0.0.0",
  includeClientSideAssets: true,
  skipFeatureDeployment: true,
  features: [{ id: "ffffffff-1111-2222-3333-444444444444", title: "Demo", description: "Demo" }],
};

const manifest: ComponentManifest = {
  id: "11111111-2222-3333-4444-555555555555",
  alias: "DemoWebPart",
  componentType: "WebPart",
  version: "1.2.3",
  manifestVersion: 2,
  requiresCustomScript: false,
  preconfiguredEntries: [{ title: { default: "Demo" } }],
  loaderConfig: {
    internalModuleBaseUrls: ["https://localhost:4321/dist/"],
    entryModuleId: "demo-web-part",
    scriptResources: { "demo-web-part": { type: "path", path: "demo-web-part_abc.js" } },
  },
};

const assets = [{ name: "demo-web-part_abc.js", data: new TextEncoder().encode("define()") }];

const pack = (overrides: Partial<Solution> = {}) =>
  unzipSync(
    createSppkg({ solution: { ...solution, ...overrides }, manifests: [manifest], assets }),
  );

const featureDir = solution.features![0]!.id;

describe("createSppkg", () => {
  it("contains the OPC parts SharePoint reads", () => {
    // Anything missing here makes the app catalog reject the upload, and its
    // error message does not say which part was wrong.
    expect(Object.keys(pack()).toSorted()).toEqual(
      [
        "[Content_Types].xml",
        "_rels/.rels",
        "AppManifest.xml",
        "_rels/AppManifest.xml.rels",
        `feature_${featureDir}.xml`,
        `feature_${featureDir}.xml.config.xml`,
        `_rels/feature_${featureDir}.xml.rels`,
        `${featureDir}/WebPart_${manifest.id}.xml`,
        "ClientSideAssets.xml",
        "ClientSideAssets.xml.config.xml",
        "_rels/ClientSideAssets.xml.rels",
        "ClientSideAssets/demo-web-part_abc.js",
      ].toSorted(),
    );
  });

  it("embeds the component manifest as valid JSON in the element XML", () => {
    // The manifest exists nowhere else in the package — no loose .json file.
    const xml = strFromU8(pack()[`${featureDir}/WebPart_${manifest.id}.xml`]!);
    const attribute = /ComponentManifest="([^"]*)"/.exec(xml)![1]!;
    const embedded = JSON.parse(
      attribute.replaceAll("&quot;", '"').replaceAll("&amp;", "&").replaceAll("&apos;", "'"),
    );

    expect(embedded.id).toBe(manifest.id);
    expect(embedded.loaderConfig.entryModuleId).toBe("demo-web-part");
    // Installation gate, stripped by the packer and mirrored as an attribute.
    expect(embedded).not.toHaveProperty("requiresCustomScript");
  });

  it("rewrites the base URL to the exact string the server matches on", () => {
    const xml = strFromU8(pack()[`${featureDir}/WebPart_${manifest.id}.xml`]!);
    expect(xml).toContain("HTTPS://SPCLIENTSIDEASSETLIBRARY/");
    expect(xml).not.toContain("localhost:4321");
  });

  it("keeps the CDN base URL when assets are not embedded", () => {
    const files = pack({ includeClientSideAssets: false });
    expect(Object.keys(files)).not.toContain("ClientSideAssets/demo-web-part_abc.js");
    const xml = strFromU8(files[`${featureDir}/WebPart_${manifest.id}.xml`]!);
    expect(xml).not.toContain("SPCLIENTSIDEASSETLIBRARY");
  });

  it("declares a content type for every extension in the package", () => {
    // OPC requires it, and a missing entry fails at upload rather than install.
    const files = pack();
    const types = strFromU8(files["[Content_Types].xml"]!);
    const extensions = new Set(
      Object.keys(files)
        .map((name) => name.slice(name.lastIndexOf(".") + 1).toLowerCase())
        .filter((ext) => ext && !ext.includes("/")),
    );
    for (const extension of extensions) {
      expect(types).toContain(`Extension="${extension}"`);
    }
  });

  it("provisions web parts into the gallery", () => {
    const xml = strFromU8(pack()[`${featureDir}/WebPart_${manifest.id}.xml`]!);
    expect(xml).toContain(`Url="_catalogs/wp"`);
    expect(xml).toContain(`List="113"`);
  });

  it("is byte-reproducible", () => {
    // Heft randomises its AppPartConfig GUIDs and stamps real timestamps, so
    // its packages differ on every build. Ours must not, otherwise "did the
    // output change?" is unanswerable.
    const first = createSppkg({ solution, manifests: [manifest], assets });
    const second = createSppkg({ solution, manifests: [manifest], assets });
    expect(Buffer.from(first).equals(Buffer.from(second))).toBe(true);
  });
});
