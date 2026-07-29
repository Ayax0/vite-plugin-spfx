import { describe, expect, it } from "vitest";
import { resolveConfig } from "vite";
import { spfx } from "../src/index.ts";
import { createFixture } from "./_fixture.ts";

/**
 * Vite merges a plugin's `config()` result OVER the user's, so a default returned
 * unconditionally cannot be overridden. The plugin therefore reads the user
 * config and only fills gaps — these tests hold that behaviour in place.
 */
const resolve = (root: string, userConfig: Record<string, unknown> = {}, mode = "production") =>
  resolveConfig(
    {
      root,
      configFile: false,
      logLevel: "silent",
      plugins: [
        spfx({
          bundleName: "demo",
          entry: "src/entry.ts",
          manifest: "src/entry.manifest.json",
        }),
      ],
      ...userConfig,
    },
    "build",
    mode,
  );

describe("build defaults", () => {
  it("targets a browser level SharePoint serves itself", async () => {
    const config = await resolve(createFixture());
    expect(config.build.target).toBe("es2017");
  });

  it("minifies for production and leaves development readable", async () => {
    expect((await resolve(createFixture())).build.minify).toBeTruthy();
    expect((await resolve(createFixture(), {}, "development")).build.minify).toBe(false);
  });

  it("emits source maps while developing, none for release — like Heft", async () => {
    expect((await resolve(createFixture())).build.sourcemap).toBe(false);
    expect((await resolve(createFixture(), {}, "development")).build.sourcemap).toBe(true);
  });

  it("lets the project override every default", async () => {
    const config = await resolve(createFixture(), {
      build: { target: "es2022", minify: false, sourcemap: true },
    });
    expect(config.build.target).toBe("es2022");
    expect(config.build.minify).toBe(false);
    expect(config.build.sourcemap).toBe(true);
  });

  it("keeps the settings the loader depends on", async () => {
    // These are not defaults: CSS must not be split out, and the entry has to
    // keep its exports or the web part class is tree-shaken away.
    const config = await resolve(createFixture(), {
      build: { cssCodeSplit: true },
    });
    expect(config.build.cssCodeSplit).toBe(false);
    expect(config.build.rollupOptions.preserveEntrySignatures).toBe("strict");
  });
});
