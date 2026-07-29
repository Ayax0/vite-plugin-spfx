import fs from "node:fs";
import path from "node:path";
import { readJsonc } from "../_utils.ts";
import type { ComponentManifest } from "../types.ts";

/**
 * The contract SharePoint's debug flow expects from a local server.
 * Reverse-engineered from @microsoft/spfx-heft-plugins.
 */

/** `config/serve.json` in the stock SPFx shape, so migrating projects keep theirs. */
export interface ServeConfig {
  port: number;
  https: boolean;
  /** May contain `{tenantDomain}`, filled from SPFX_SERVE_TENANT_DOMAIN. */
  initialPage?: string;
}

export function readServeConfig(root: string): ServeConfig {
  const file = path.join(root, "config", "serve.json");
  const config = fs.existsSync(file)
    ? readJsonc<
        Partial<ServeConfig> & { serveConfigurations?: Record<string, { pageUrl?: string }> }
      >(file)
    : {};

  return {
    port: config.port ?? 4321,
    https: config.https ?? true,
    initialPage: config.initialPage ?? config.serveConfigurations?.["default"]?.pageUrl,
  };
}

/** The SPFx >= 1.21 location. SharePoint still probes the older `/temp/` one. */
export const MANIFESTS_PATH = "/temp/build/manifests.js";
export const LEGACY_MANIFESTS_PATH = "/temp/manifests.js";

/**
 * `debug=true` is required by the client gate, and the host of
 * `debugManifestsFile` must be exactly localhost or 127.0.0.1. `noredir=true`
 * sends loader errors to the console instead of SharePoint's error page.
 */
export function debugQuery(manifestsUrl: string): string {
  const params = new URLSearchParams({
    debugManifestsFile: manifestsUrl,
    debug: "true",
    noredir: "true",
  });
  return `?${params.toString()}`;
}

/**
 * Fill `{tenantDomain}` the way SPFx does.
 *
 * The tenant is per-developer data — everyone on a team works against a
 * different one — so it comes from the environment rather than committed
 * config. Returns undefined when the placeholder cannot be filled, so callers
 * print instructions instead of a broken link.
 */
export function resolveInitialPage(
  initialPage: string | undefined,
  tenantDomain: string | undefined,
): string | undefined {
  if (!initialPage) return undefined;
  if (!initialPage.includes("{tenantDomain}")) return initialPage;
  return tenantDomain ? initialPage.replace("{tenantDomain}", tenantDomain) : undefined;
}

/**
 * The workbench and the Debug Toolbar load this as an AMD module and call
 * `getManifests()`. Only the project's own manifests are needed — SharePoint
 * filters framework manifests out for third parties anyway.
 */
export function manifestsModule(manifests: ComponentManifest[]): string {
  return (
    `var manifests = ${JSON.stringify(manifests, null, 2)};\n` +
    `var api = { _metadata: undefined, getManifests: function () { return manifests; } };\n` +
    `self.debugManifests = api;\n` +
    `if (typeof define === "function" && define.amd) { define([], function () { return api; }); }\n`
  );
}
