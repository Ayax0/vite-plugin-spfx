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

const MANIFEST_FILE = /^[^.\\/]+\.manifest\.json$/;

interface DiscoveredManifest {
  id?: string;
  version?: string;
  /** Present on "assembly" manifests, pointing at the real component. */
  rootComponentId?: string;
}

/**
 * Find every dependency SharePoint already serves, by reading the
 * `dist/<guid>.manifest.json` each SPFx package ships. Discovering this instead
 * of hardcoding it is what lets the plugin survive SPFx version bumps.
 */
export function discoverSpfxExternals(root: string): Map<string, SpfxComponent> {
  const resolve = createRequire(path.join(root, "package.json")).resolve;
  const pkg = readJsonc<{
    dependencies?: Record<string, string>;
    devDependencies?: Record<string, string>;
  }>(path.join(root, "package.json"));

  const found = new Map<string, SpfxComponent>();

  for (const name of Object.keys({ ...pkg.dependencies, ...pkg.devDependencies })) {
    let dir: string;
    try {
      dir = path.dirname(resolve(`${name}/package.json`));
    } catch {
      continue; // types-only or otherwise unresolvable
    }
    const installedVersion = () =>
      cleanVersion(readJsonc<{ version: string }>(path.join(dir, "package.json")).version);

    const hardcoded = NON_STANDARD_COMPONENTS[name];
    if (hardcoded) {
      found.set(name, { id: hardcoded, version: installedVersion() });
      continue;
    }

    const distDir = path.join(dir, "dist");
    if (!fs.existsSync(distDir)) continue;

    const manifests = fs
      .readdirSync(distDir)
      .filter((file) => MANIFEST_FILE.test(file))
      .flatMap((file) => {
        try {
          return [readJsonc<DiscoveredManifest>(path.join(distDir, file))];
        } catch {
          return [];
        }
      });

    // A package may ship a component plus its assembly; only the component —
    // the one without a rootComponentId — is importable.
    const components = manifests.filter((m) => m.id && !m.rootComponentId);
    if (components.length !== 1) continue;

    const manifest = components[0]!;
    found.set(name, {
      id: manifest.id!,
      version:
        !manifest.version || manifest.version === "*"
          ? installedVersion()
          : cleanVersion(manifest.version),
    });
  }

  return found;
}
