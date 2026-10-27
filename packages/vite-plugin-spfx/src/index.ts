import fs from "node:fs";
import path from "node:path";
import type { Plugin, Rolldown } from "vite";

import { cleanVersion, readJsonc } from "./_utils.ts";
import { spfxServe } from "./serve.ts";
import { readServeConfig } from "./spfx/debug.ts";
import { createComponentLookup } from "./spfx/externals.ts";
import { buildScriptResources, collectLocalizedResource } from "./spfx/manifest.ts";
import { createSppkg } from "./spfx/solution.ts";
import { esmToAmd } from "./transform/amd.ts";
import {
  rewriteThemeTokens,
  spfxThemeTokens,
  STYLE_TOKEN_PATTERN,
  styleInjector,
} from "./transform/styles.ts";
import type {
  ComponentManifest,
  PackageSolutionConfig,
  ScriptResource,
  SpfxComponent,
  SpfxOptions,
} from "./types.ts";

export { spfxServe } from "./serve.ts";
export { createSppkg } from "./spfx/solution.ts";
export { esmToAmd } from "./transform/amd.ts";
export { rewriteThemeTokens } from "./transform/styles.ts";
export type * from "./types.ts";

/**
 * Turn a Vite project into a SharePoint Framework component.
 *
 * `vite build` emits the named-AMD bundle, the component manifest and a
 * deployable `.sppkg`; `vite dev` serves them to SharePoint's debug flow.
 *
 * The order of the build hooks is load-bearing, not incidental: minification
 * runs between `renderChunk` and `generateBundle`, so anything that inspects or
 * rewrites the final code has to sit on the right side of it.
 */
export function spfx(options: SpfxOptions): Plugin[] {
  return [spfxBuild(options), spfxServe(options)];
}

function spfxBuild(options: SpfxOptions): Plugin {
  const {
    bundleName,
    entry,
    manifest: manifestPath,
    localizedResources = {},
    cdnBasePath = "",
  } = options;

  let root: string;
  let outDir: string;
  let production = false;
  let serveBaseUrl: string;
  let findComponent: ReturnType<typeof createComponentLookup>;
  /** The components the bundle imports, collected as Rolldown meets them. */
  let externals: Map<string, SpfxComponent>;
  let manifest: ComponentManifest;
  let amdId: string;
  /** Dependency order established in renderChunk, consumed in generateBundle. */
  let deps: string[] = [];
  let releaseManifest: ComponentManifest | undefined;

  return {
    name: "vite-plugin-spfx",
    enforce: "post",
    apply: "build",

    config(userConfig, env) {
      root = path.resolve(userConfig.root ?? process.cwd());
      production = env.mode === "production";
      findComponent = createComponentLookup(root);
      externals = new Map();

      // The debug manifest has to point at wherever the dev server will listen,
      // so derive it from the same config the server reads.
      const serve = readServeConfig(root);
      const port = userConfig.server?.port ?? serve.port;
      serveBaseUrl =
        options.serveBaseUrl ?? `${serve.https ? "https" : "http"}://localhost:${port}/dist/`;

      manifest = readJsonc<ComponentManifest>(path.resolve(root, manifestPath));
      delete manifest.$schema;
      if (!manifest.version || manifest.version === "*") {
        manifest.version = cleanVersion(
          readJsonc<{ version: string }>(path.join(root, "package.json")).version,
        );
      }
      amdId = `${manifest.id}_${manifest.version}`;

      const localizedNames = new Set(Object.keys(localizedResources));

      // Vite merges a plugin's config OVER the user's, so returning a value
      // unconditionally would make it impossible to override. Read what the
      // project set and only fill the gaps.
      const userBuild = userConfig.build ?? {};

      return {
        build: {
          // --- Required by the SPFx loader; overriding these breaks the output.
          //
          // Deliberately NOT build.lib: lib mode always inlines assets as data
          // URIs because it cannot know the public path. We do know it at
          // runtime, and assets must be real files to reach ClientSideAssets/.
          cssCodeSplit: false,
          modulePreload: false,
          rollupOptions: {
            input: path.resolve(root, entry),
            // Decided per import, so a component stays external whichever
            // package imports it — see createComponentLookup.
            external: (id: string, importer: string | undefined) => {
              if (localizedNames.has(id)) return true;
              const known = externals.get(id);
              // A virtual importer resolves nothing itself; reuse what is known.
              const component = findComponent(id, importer) ?? known;
              if (!component) return false;
              // Without a copy of its own in the project, two dependencies can
              // bring different versions. Keeping whichever Rolldown meets first
              // would make the manifest change from build to build.
              if (known && known.version !== component.version) {
                throw new Error(
                  `[spfx] "${id}" is installed as both ${known.version} and ` +
                    `${component.version}. Add it to package.json to pick one.`,
                );
              }
              externals.set(id, component);
              return true;
            },
            // Without this the entry counts as an application root and its
            // exports — the web part class included — are tree-shaken away.
            preserveEntrySignatures: "strict",
            output: {
              format: "es",
              // A content hash in production keeps the app catalog from serving
              // a stale bundle after an upgrade; serve mode stays stable so
              // already-deployed debug manifests keep resolving.
              entryFileNames: production ? `${bundleName}_[hash].js` : `${bundleName}.js`,
              assetFileNames: "[name]-[hash][extname]",
              // One self-contained file per component: the SPFx loader offers no
              // usable AMD require for extra chunks (its baseUrl is a poison URL).
              codeSplitting: false,
            },
          },

          // --- Defaults, chosen for SharePoint but yours to change.
          //
          // SharePoint Online runs on evergreen browsers; ES2017 is comfortably
          // within what it serves itself.
          target: userBuild.target ?? "es2017",
          // Mirrors Heft: readable output while developing, minified to ship.
          minify: userBuild.minify ?? production,
          // Also mirrors Heft, which emits no source maps for release builds.
          sourcemap: userBuild.sourcemap ?? !production,
        },
        // Must run before the CSS minifier drops the themed declarations.
        css: { postcss: { plugins: [spfxThemeTokens()] } },
        // Asset URLs resolve against the script's own location, exactly like
        // SPFx's SetPublicPathCurrentScriptPlugin. __spfx_base is set by the
        // AMD header in amd.ts.
        experimental: {
          renderBuiltUrl(filename: string, { hostType }: { hostType: string }) {
            return hostType === "js"
              ? { runtime: `__spfx_base + ${JSON.stringify(filename)}` }
              : { relative: true as const };
          },
        },
      };
    },

    configResolved(resolved) {
      outDir = path.resolve(resolved.root, resolved.build.outDir);
    },

    /**
     * The ESM -> AMD rewrite belongs here: renderChunk is the documented place to
     * transform emitted code, and the content hash is computed from its result.
     * Doing it in generateBundle would leave the hash blind to the wrapper, so a
     * plugin-side fix would ship under the old name and stay cached.
     */
    async renderChunk(code, chunk) {
      if (!chunk.isEntry) return null;
      const result = await esmToAmd(code, { amdId, prelude: [styleInjector(amdId)] });
      deps = result.deps;
      return { code: result.code, map: result.map };
    },

    async generateBundle(_options, bundle) {
      // SPFx manifests have no CSS resource type, so styles must travel inside
      // the JS. generateBundle is the only point where Vite's finished CSS exists.
      let styles = "";
      for (const [name, output] of Object.entries(bundle)) {
        if (output.type !== "asset" || !name.endsWith(".css")) continue;
        styles += rewriteThemeTokens(String(output.source));
        delete bundle[name];
      }

      const entryFileName = Object.keys(bundle).find((name) => {
        const output = bundle[name];
        return output?.type === "chunk" && output.isEntry;
      });
      if (!entryFileName) throw new Error("[spfx] no entry chunk found");
      const chunk = bundle[entryFileName] as Rolldown.OutputChunk;
      if (!STYLE_TOKEN_PATTERN.test(chunk.code)) {
        throw new Error(
          "[spfx] the style placeholder did not survive rendering — styles would be lost",
        );
      }
      chunk.code = chunk.code.replace(STYLE_TOKEN_PATTERN, JSON.stringify(styles));

      const localized: Record<string, ScriptResource> = {};
      for (const [name, pattern] of Object.entries(localizedResources)) {
        const { resource, files } = collectLocalizedResource(root, name, pattern, production);
        for (const file of files) {
          this.emitFile({ type: "asset", fileName: file.fileName, source: file.source });
        }
        localized[name] = resource;
      }

      const debug = structuredClone(manifest);
      debug.loaderConfig = {
        internalModuleBaseUrls: [serveBaseUrl],
        entryModuleId: bundleName,
        scriptResources: buildScriptResources(
          bundleName,
          entryFileName,
          deps,
          externals,
          localized,
        ),
      };
      this.emitFile({
        type: "asset",
        fileName: `${debug.id}.manifest.json`,
        source: JSON.stringify(debug, null, 2),
      });

      // The release variant differs only in internalModuleBaseUrls.
      releaseManifest = structuredClone(debug);
      releaseManifest.loaderConfig!.internalModuleBaseUrls = cdnBasePath ? [cdnBasePath] : [];

      this.info(`${bundleName} -> define(${JSON.stringify(amdId)}, [${deps.join(", ")}])`);
    },

    writeBundle() {
      if (!releaseManifest) return;
      const release = releaseManifest;
      releaseManifest = undefined;

      // Read the shippable files from disk, not from the bundle object: assets
      // emitted via emitFile only appear there after generateBundle, and Rolldown
      // keeps deleted entries enumerable within the same hook. Stylesheets are
      // excluded because they are inlined into the JS — dropping them from the
      // bundle is best-effort, so a stray file must never reach the package.
      const assets = fs
        .readdirSync(outDir)
        .filter((file) => !/\.(manifest\.json|map|css)$/.test(file))
        .map((file) => ({ name: file, data: fs.readFileSync(path.join(outDir, file)) }));

      const releaseDir = path.join(root, "release");
      fs.mkdirSync(path.join(releaseDir, "manifests"), { recursive: true });
      fs.mkdirSync(path.join(releaseDir, "assets"), { recursive: true });
      fs.writeFileSync(
        path.join(releaseDir, "manifests", `${release.id}.manifest.json`),
        JSON.stringify(release, null, 2),
      );
      for (const asset of assets) {
        fs.writeFileSync(path.join(releaseDir, "assets", asset.name), asset.data);
      }

      if (!production) return; // Heft also refuses to package non-production builds

      const configPath = path.join(root, "config", "package-solution.json");
      if (!fs.existsSync(configPath)) return;
      const config = readJsonc<PackageSolutionConfig>(configPath);

      // The teams/ folder (Teams app icons) always ships, if present.
      const teamsDir = path.join(root, "teams");
      if (fs.existsSync(teamsDir)) {
        for (const file of fs.readdirSync(teamsDir)) {
          assets.push({ name: file, data: fs.readFileSync(path.join(teamsDir, file)) });
        }
      }

      const zip = createSppkg({ solution: config.solution, manifests: [release], assets });
      const target = path.join(
        root,
        "sharepoint",
        config.paths?.zippedPackage ?? "solution/ClientSolution.sppkg",
      );
      fs.mkdirSync(path.dirname(target), { recursive: true });
      fs.writeFileSync(target, zip);
      this.info(`sppkg ${path.relative(root, target)} (${(zip.length / 1024).toFixed(1)} kB)`);
    },
  };
}
