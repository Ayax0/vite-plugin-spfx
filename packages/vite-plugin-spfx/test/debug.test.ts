import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import {
  debugQuery,
  manifestsModule,
  readServeConfig,
  resolveInitialPage,
} from "../src/spfx/debug.ts";
import type { ComponentManifest } from "../src/types.ts";

function project(serveJson?: unknown): string {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "spfx-serve-"));
  if (serveJson !== undefined) {
    fs.mkdirSync(path.join(root, "config"));
    fs.writeFileSync(path.join(root, "config/serve.json"), JSON.stringify(serveJson));
  }
  return root;
}

describe("readServeConfig", () => {
  it("falls back to the SPFx defaults without a serve.json", () => {
    expect(readServeConfig(project())).toEqual({
      port: 4321,
      https: true,
      initialPage: undefined,
    });
  });

  it("reads the stock SPFx file, so migrating projects keep theirs", () => {
    const root = project({
      port: 5432,
      https: false,
      initialPage: "https://{tenantDomain}/_layouts/workbench.aspx",
    });
    expect(readServeConfig(root)).toEqual({
      port: 5432,
      https: false,
      initialPage: "https://{tenantDomain}/_layouts/workbench.aspx",
    });
  });

  it("accepts a serveConfigurations entry as the page", () => {
    const root = project({ serveConfigurations: { default: { pageUrl: "https://x/page.aspx" } } });
    expect(readServeConfig(root).initialPage).toBe("https://x/page.aspx");
  });
});

describe("resolveInitialPage", () => {
  it("substitutes the tenant", () => {
    expect(
      resolveInitialPage(
        "https://{tenantDomain}/_layouts/workbench.aspx",
        "contoso.sharepoint.com",
      ),
    ).toBe("https://contoso.sharepoint.com/_layouts/workbench.aspx");
  });

  it("gives up rather than printing a broken link", () => {
    expect(resolveInitialPage("https://{tenantDomain}/x.aspx", undefined)).toBeUndefined();
  });

  it("passes a fully qualified page through, tenant or not", () => {
    expect(resolveInitialPage("https://contoso.sharepoint.com/x.aspx", undefined)).toBe(
      "https://contoso.sharepoint.com/x.aspx",
    );
  });

  it("has nothing to resolve without a page", () => {
    expect(resolveInitialPage(undefined, "contoso.sharepoint.com")).toBeUndefined();
  });
});

describe("debugQuery", () => {
  it("carries the three parameters the client gate needs", () => {
    const params = new URLSearchParams(
      debugQuery("https://localhost:4321/temp/build/manifests.js").slice(1),
    );
    // debug=true is required by _is3PDebugEnabled; noredir sends loader errors
    // to the console instead of SharePoint's error page.
    expect(params.get("debug")).toBe("true");
    expect(params.get("noredir")).toBe("true");
    expect(params.get("debugManifestsFile")).toBe("https://localhost:4321/temp/build/manifests.js");
  });
});

describe("manifestsModule", () => {
  const manifest = { id: "abc", alias: "Demo" } as unknown as ComponentManifest;

  it("exposes getManifests both globally and as an anonymous AMD module", () => {
    // sp-loader script-loads this file and calls getManifests() on the result.
    const code = manifestsModule([manifest]);
    expect(code).toContain("self.debugManifests = api");
    expect(code).toContain("define([], function () { return api; })");
  });

  it("is evaluable and returns what it was given", () => {
    const code = manifestsModule([manifest]);
    const self: Record<string, any> = {};
    new Function("self", "define", code)(self, undefined);
    expect(self["debugManifests"].getManifests()).toEqual([manifest]);
  });
});
