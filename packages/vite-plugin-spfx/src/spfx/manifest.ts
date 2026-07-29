import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import type { ScriptResource, SpfxComponent } from "../types.ts";

export interface LocalizedResource {
  resource: ScriptResource;
  files: Array<{ fileName: string; source: Buffer }>;
}

const upperFirst = (value: string): string => value.charAt(0).toUpperCase() + value.slice(1);

/**
 * "en-us" -> "en-US", "sr-cyrl-rs" -> "sr-Cyrl-RS", pseudo-locales stay lowercase.
 *
 * SharePoint matches the page language against these keys, so the casing is not
 * cosmetic. Ported from SPFx's own normalizeLocaleName so both toolchains agree.
 */
export function normalizeLocaleName(locale: string): string {
  const tokens = locale.toLowerCase().split("-");

  if (tokens.length === 2) {
    const last = tokens[1]!;
    if (last.length === 2) tokens[1] = last.toUpperCase();
    else if (!["ploc", "plocm", "ploca"].includes(last)) tokens[1] = upperFirst(last);
  } else if (tokens.length === 3) {
    const last = tokens[2]!;
    if (last.length === 2) {
      tokens[2] = last.toUpperCase();
      tokens[1] = upperFirst(tokens[1]!);
    } else {
      tokens[2] = upperFirst(last);
      tokens[1] = tokens[1]!.toUpperCase();
    }
  }

  return tokens.join("-");
}

/**
 * Expand a `.../loc/{locale}.js` pattern into one standalone AMD file per locale.
 *
 * The loc files are authored as `define([], function () { return {...} })` and
 * are copied verbatim, exactly as the SPFx toolchain does — they are already
 * valid AMD modules, so compiling them would only risk changing semantics.
 */
export function collectLocalizedResource(
  root: string,
  name: string,
  pattern: string,
  hashed: boolean,
): LocalizedResource {
  const dir = path.resolve(root, path.dirname(pattern));
  const base = path.basename(pattern);
  const matcher = new RegExp(`^${base.replace(/\./g, "\\.").replace("{locale}", "(.+)")}$`);

  const files: LocalizedResource["files"] = [];
  const paths: Record<string, string> = {};

  for (const file of fs.readdirSync(dir)) {
    const match = matcher.exec(file);
    if (!match) continue;
    const locale = match[1]!;
    const source = fs.readFileSync(path.join(dir, file));
    const suffix = hashed ? `_${createHash("md5").update(source).digest("hex")}` : "";
    const fileName = `${name}_${locale}${suffix}.js`;
    files.push({ fileName, source });
    // File name keeps the locale as authored; the manifest key is normalised,
    // because that is what SharePoint matches the page language against.
    paths[normalizeLocaleName(locale)] = fileName;
  }

  const locales = Object.keys(paths);
  if (locales.length === 0) throw new Error(`[spfx] no locale files matched ${pattern}`);

  return {
    files,
    resource:
      locales.length === 1
        ? { type: "path", path: paths[locales[0]!]! }
        : { type: "localizedPath", defaultPath: paths["en-US"] ?? paths[locales[0]!]!, paths },
  };
}

/**
 * Build `loaderConfig.scriptResources`.
 *
 * The keys must be exactly the names in the `define()` dependency array, and in
 * the same order: the loader maps them positionally onto the factory arguments.
 */
export function buildScriptResources(
  bundleName: string,
  entryFileName: string,
  deps: string[],
  externals: Map<string, SpfxComponent>,
  localized: Record<string, ScriptResource>,
): Record<string, ScriptResource> {
  const resources: Record<string, ScriptResource> = {
    [bundleName]: { type: "path", path: entryFileName },
  };

  for (const dep of deps) {
    const component = externals.get(dep);
    if (component) {
      resources[dep] = { type: "component", id: component.id, version: component.version };
      continue;
    }
    const strings = localized[dep];
    if (strings) {
      resources[dep] = strings;
      continue;
    }
    throw new Error(
      `[spfx] external "${dep}" has no scriptResources entry — the SharePoint loader ` +
        `would fail to resolve it. Add it as a dependency or bundle it.`,
    );
  }

  return resources;
}
