import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { cleanVersion, readJsonc } from "../_utils.ts";
import type { SpfxComponent } from "../types.ts";

/**
 * Packages that ARE SharePoint components at runtime but ship no discoverable
 * `dist/<guid>.manifest.json`. SPFx hardcodes exactly these three
 * (spfx-heft-plugins → spfxManifests/utilities/ExternalsProcessor.js).
 *
 * Every other component id is read from the installed package, so this map is
 * the plugin's only standing bet on SPFx internals — exported so the baseline
 * test can hold it against Microsoft's own build output.
 */
export const NON_STANDARD_COMPONENTS: Record<string, string> = {
  react: "0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d",
  "react-dom": "aa0a46ec-1505-43cd-a44a-93f3a5aa460a",
  "@microsoft/microsoft-graph-client": "32b9363b-09ff-4959-929a-cab2663fb223",
};

const MANIFEST_FILE = /^([^.\\/]+)\.manifest\.json$/;

/** A whole package name: SharePoint serves a component as one unit, never a subpath. */
const PACKAGE_NAME = /^(?:@[\w~-][\w.~-]*\/)?[\w~-][\w.~-]*$/;

interface DiscoveredManifest {
  id?: string;
  version?: string;
  /** Present on "assembly" manifests, pointing at the real component. */
  rootComponentId?: string;
}

/** The component the package installed at `dir` represents, if it is one. */
function readComponent(name: string, dir: string): SpfxComponent | undefined {
  const installedVersion = () =>
    cleanVersion(readJsonc<{ version: string }>(path.join(dir, "package.json")).version);

  const hardcoded = NON_STANDARD_COMPONENTS[name];
  if (hardcoded) return { id: hardcoded, version: installedVersion() };

  const distDir = path.join(dir, "dist");
  if (!fs.existsSync(distDir)) return undefined;

  const manifests = fs.readdirSync(distDir).flatMap((file) => {
    const id = MANIFEST_FILE.exec(file)?.[1];
    if (!id) return [];
    try {
      const manifest = readJsonc<DiscoveredManifest>(path.join(distDir, file));
      // SPFx names every manifest after its id. Any other *.manifest.json belongs
      // to a different tool, and every imported package is checked, not only
      // SPFx ones — mistaking a library for a component would break it at runtime.
      return manifest.id === id ? [manifest] : [];
    } catch {
      return [];
    }
  });

  // A package may ship a component plus its assembly; only the component —
  // the one without a rootComponentId — is importable.
  const [manifest, ...others] = manifests.filter((m) => m.id && !m.rootComponentId);
  if (!manifest?.id || others.length > 0) return undefined;

  return {
    id: manifest.id,
    version:
      !manifest.version || manifest.version === "*"
        ? installedVersion()
        : cleanVersion(manifest.version),
  };
}

/**
 * Decide import by import whether a specifier names a package SharePoint
 * already serves, by reading the `dist/<guid>.manifest.json` each SPFx package
 * ships. Discovering this instead of hardcoding it is what lets the plugin
 * survive SPFx version bumps.
 *
 * Deliberately not a scan of package.json: components reach the bundle as
 * hoisted transitive dependencies, from a dependency's own node_modules, or
 * through a library — and one that slips through drags SharePoint's internals
 * into the bundle. Heft follows dependencies only through packages that are
 * components themselves, so it misses the last case; this lookup does not.
 *
 * The project's own copy wins over one nested in a dependency: scriptResources
 * holds one version per name, and the project's is the one it asked for.
 */
export function createComponentLookup(
  root: string,
): (name: string, importer?: string) => SpfxComponent | undefined {
  const cache = new Map<string, SpfxComponent | undefined>();

  const find = (name: string, from: string): SpfxComponent | undefined => {
    let dir: string;
    try {
      dir = path.dirname(createRequire(from).resolve(`${name}/package.json`));
    } catch {
      return undefined; // not installed where `from` can see it, or a virtual importer
    }
    if (!cache.has(dir)) cache.set(dir, readComponent(name, dir));
    return cache.get(dir);
  };

  return (name, importer) => {
    if (!PACKAGE_NAME.test(name)) return undefined;
    return (
      find(name, path.join(root, "package.json")) ?? (importer ? find(name, importer) : undefined)
    );
  };
}
