import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { build } from "vite";
import { spfx } from "../src/index.ts";
import type { SpfxOptions } from "../src/types.ts";

/**
 * Fake SPFx runtime packages, written into the fixture's node_modules.
 *
 * The plugin identifies SharePoint components purely by the
 * `dist/<guid>.manifest.json` every SPFx package ships, so a two-file stand-in
 * exercises the real discovery path. Using the actual @microsoft/sp-* packages
 * would make the tests slow, huge and dependent on whichever SPFx version
 * happens to be installed — the opposite of a stable baseline.
 */
export interface FakeComponent {
  name: string;
  id: string;
  version: string;
  /** Emits a second manifest carrying rootComponentId, like sp-loader does. */
  withAssembly?: boolean;
  /** Omit the dist manifest entirely — the package must then be bundled, not externalised. */
  plain?: boolean;
}

export interface FixtureOptions {
  components?: FakeComponent[];
  /** Import statements and body of the entry module. */
  entry?: string;
  manifest?: Record<string, unknown>;
  styles?: string;
  locales?: Record<string, string>;
  packageSolution?: Record<string, unknown>;
  version?: string;
}

const COMPONENT_ID = "11111111-2222-3333-4444-555555555555";

export const DEFAULT_MANIFEST = {
  id: COMPONENT_ID,
  alias: "DemoWebPart",
  componentType: "WebPart",
  version: "*",
  manifestVersion: 2,
  requiresCustomScript: false,
  supportedHosts: ["SharePointWebPart"],
  preconfiguredEntries: [
    { group: { default: "Advanced" }, title: { default: "Demo" }, properties: {} },
  ],
};

export function createFixture(options: FixtureOptions = {}): string {
  const {
    components = [],
    entry = `export default class DemoWebPart {}`,
    manifest = DEFAULT_MANIFEST,
    styles,
    locales,
    packageSolution,
    version = "1.2.3",
  } = options;

  const root = fs.mkdtempSync(path.join(os.tmpdir(), "spfx-fixture-"));
  const write = (relative: string, content: string) => {
    const file = path.join(root, relative);
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, content);
  };

  write(
    "package.json",
    JSON.stringify({
      name: "fixture",
      version,
      dependencies: Object.fromEntries(components.map((c) => [c.name, "1.0.0"])),
    }),
  );

  for (const component of components) {
    const dir = `node_modules/${component.name}`;
    write(
      `${dir}/package.json`,
      JSON.stringify({ name: component.name, version: component.version }),
    );
    if (component.plain) continue;
    write(
      `${dir}/dist/${component.id}.manifest.json`,
      JSON.stringify({ id: component.id, version: component.version, componentType: "Library" }),
    );
    if (component.withAssembly) {
      const assemblyId = component.id.replace(/^./, "9");
      write(
        `${dir}/dist/${assemblyId}.manifest.json`,
        JSON.stringify({ id: assemblyId, version: "1.0.0", rootComponentId: component.id }),
      );
    }
  }

  write("src/entry.ts", entry);
  write("src/entry.manifest.json", JSON.stringify(manifest, null, 2));
  if (styles) write("src/styles.css", styles);
  for (const [locale, content] of Object.entries(locales ?? {})) {
    write(`src/loc/${locale}.js`, content);
  }
  if (packageSolution) {
    write("config/package-solution.json", JSON.stringify(packageSolution, null, 2));
  }

  return root;
}

/** Run a real production build over a fixture — the only way to test what ships. */
export async function buildFixture(
  root: string,
  options: Partial<SpfxOptions> = {},
): Promise<string> {
  await build({
    root,
    mode: "production",
    logLevel: "silent",
    configFile: false,
    plugins: [
      spfx({
        bundleName: "demo-web-part",
        entry: "src/entry.ts",
        manifest: "src/entry.manifest.json",
        ...options,
      }),
    ],
  });
  return path.join(root, "dist");
}

export const readDist = (dist: string): string[] => fs.readdirSync(dist).sort();

export function readEntry(dist: string): { fileName: string; code: string } {
  const fileName = readDist(dist).find(
    (file) => file.startsWith("demo-web-part") && file.endsWith(".js"),
  )!;
  return { fileName, code: fs.readFileSync(path.join(dist, fileName), "utf8") };
}

export function readManifest(dist: string): Record<string, any> {
  const file = readDist(dist).find((name) => name.endsWith(".manifest.json"))!;
  return JSON.parse(fs.readFileSync(path.join(dist, file), "utf8"));
}

/**
 * Read the `define(id, deps, factory)` header out of a built bundle.
 *
 * Quote-agnostic on purpose: the minifier rewrites string literals to
 * backticks, so assertions that spell out a quote character break on the
 * production build only — the exact failure mode these tests exist to catch.
 */
export function parseDefine(code: string): { id: string; deps: string[] } {
  const match = /\bdefine\(\s*(["'`])(.+?)\1\s*,\s*\[([^\]]*)\]/.exec(code);
  if (!match) throw new Error("no named define() found in the bundle");
  const list = match[3]!.trim();
  return {
    id: match[2]!,
    deps: list ? list.split(",").map((dep) => dep.trim().replace(/^["'`]|["'`]$/g, "")) : [],
  };
}
