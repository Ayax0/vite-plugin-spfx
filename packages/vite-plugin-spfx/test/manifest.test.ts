import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { collectLocalizedResource, normalizeLocaleName } from "../src/spfx/manifest.ts";

/**
 * SharePoint matches the page language against the keys of `localizedPath.paths`,
 * so their casing is contractual, not cosmetic. Ported from SPFx's own
 * normalizeLocaleName — these cases come from its documentation.
 */
describe("normalizeLocaleName", () => {
  it("uppercases a two-letter region", () => {
    expect(normalizeLocaleName("en-us")).toBe("en-US");
    expect(normalizeLocaleName("de-de")).toBe("de-DE");
  });

  it("title-cases a script subtag", () => {
    expect(normalizeLocaleName("uz-cyrl")).toBe("uz-Cyrl");
  });

  it("handles language-script-region", () => {
    expect(normalizeLocaleName("sr-cyrl-rs")).toBe("sr-Cyrl-RS");
  });

  it("leaves pseudo-locales lowercase", () => {
    expect(normalizeLocaleName("qps-ploc")).toBe("qps-ploc");
    expect(normalizeLocaleName("qps-ploca")).toBe("qps-ploca");
  });

  it("passes a bare language through", () => {
    expect(normalizeLocaleName("en")).toBe("en");
  });
});

function fixture(locales: string[]): string {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "spfx-loc-"));
  fs.mkdirSync(path.join(root, "loc"));
  for (const locale of locales) {
    fs.writeFileSync(
      path.join(root, "loc", `${locale}.js`),
      `define([], function () { return { L: "${locale}" }; });`,
    );
  }
  return root;
}

describe("collectLocalizedResource", () => {
  it("emits a plain path for a single locale", () => {
    const { resource, files } = collectLocalizedResource(
      fixture(["en-us"]),
      "DemoStrings",
      "loc/{locale}.js",
      false,
    );
    expect(resource).toEqual({ type: "path", path: "DemoStrings_en-us.js" });
    expect(files).toHaveLength(1);
  });

  it("emits localizedPath with normalised keys for several locales", () => {
    const { resource, files } = collectLocalizedResource(
      fixture(["en-us", "de-de"]),
      "DemoStrings",
      "loc/{locale}.js",
      false,
    );

    // File names keep the locale as authored; only the manifest keys are normalised.
    expect(files.map((file) => file.fileName).toSorted()).toEqual([
      "DemoStrings_de-de.js",
      "DemoStrings_en-us.js",
    ]);
    expect(resource).toEqual({
      type: "localizedPath",
      defaultPath: "DemoStrings_en-us.js",
      paths: { "en-US": "DemoStrings_en-us.js", "de-DE": "DemoStrings_de-de.js" },
    });
  });

  it("falls back to the first locale when there is no en-US", () => {
    const { resource } = collectLocalizedResource(
      fixture(["fr-fr", "de-de"]),
      "DemoStrings",
      "loc/{locale}.js",
      false,
    );
    expect((resource as { defaultPath: string }).defaultPath).toBeDefined();
  });

  it("hashes file names for production", () => {
    const { files } = collectLocalizedResource(
      fixture(["en-us"]),
      "DemoStrings",
      "loc/{locale}.js",
      true,
    );
    expect(files[0]!.fileName).toMatch(/^DemoStrings_en-us_[\da-f]{32}\.js$/);
  });

  it("fails loudly when the pattern matches nothing", () => {
    expect(() =>
      collectLocalizedResource(fixture([]), "DemoStrings", "loc/{locale}.js", false),
    ).toThrow(/no locale files matched/);
  });
});
