import fs from "node:fs";
import path from "node:path";

/** Files whose contents get placeholder substitution; anything else is copied byte for byte. */
const TEXT_FILES = /\.(json|ts|tsx|js|mjs|scss|css|vue|svelte|md|svg)$/;

/**
 * npm silently drops a file named `.gitignore` from a published tarball, so
 * templates carry it under a neutral name and it is restored on the way out.
 */
const RENAMED: Record<string, string> = { _gitignore: ".gitignore" };

export type Replacements = Record<string, string>;

/** Replace every `__TOKEN__` that has a value; unknown tokens are left untouched. */
export function substitute(content: string, replacements: Replacements): string {
  return content.replace(/__[A-Z0-9_]+__/g, (token) => replacements[token] ?? token);
}

/** Copy a template tree into `targetDir`, substituting as it goes. */
export function scaffold(
  templateDir: string,
  targetDir: string,
  replacements: Replacements,
): string[] {
  const written: string[] = [];

  const walk = (from: string, to: string): void => {
    fs.mkdirSync(to, { recursive: true });
    for (const entry of fs.readdirSync(from, { withFileTypes: true })) {
      const source = path.join(from, entry.name);
      // File and folder names carry placeholders too — the component directory
      // and the class file are named after the web part, not after the sample.
      const target = path.join(to, substitute(RENAMED[entry.name] ?? entry.name, replacements));

      if (entry.isDirectory()) {
        walk(source, target);
        continue;
      }

      if (TEXT_FILES.test(entry.name)) {
        fs.writeFileSync(target, substitute(fs.readFileSync(source, "utf8"), replacements));
      } else {
        fs.copyFileSync(source, target);
      }
      written.push(path.relative(targetDir, target));
    }
  };

  walk(templateDir, targetDir);
  return written.toSorted();
}

/**
 * Turn a directory name into an npm-safe package name.
 *
 * Also used for the solution name, which ends up in the `.sppkg` file name.
 */
export function toPackageName(input: string): string {
  const name = input
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9-~][^a-z0-9-._~]*/g, "-")
    .replace(/^[-_.]+|[-_.]+$/g, "");
  return name || "spfx-web-part";
}

export interface ComponentNames {
  /** Folder under `src/webparts/`, e.g. `teamRoster`. */
  dir: string;
  /** The framework component, e.g. `TeamRoster`. */
  component: string;
  /** Web part class and manifest alias, e.g. `TeamRosterWebPart`. */
  alias: string;
  /** Bundle key and output file name, e.g. `team-roster-web-part`. */
  bundleName: string;
  /** Bare specifier the localized strings resolve under. */
  strings: string;
}

/**
 * Identifiers the templates already import.
 *
 * A component named `React` would be declared twice in the same file as
 * `import * as React from "react"` — the build fails with a duplicate
 * identifier, which reads like a plugin bug rather than a naming clash.
 */
const RESERVED = new Set(["React", "ReactDom", "ReactDOM", "Vue", "Svelte", "Solid"]);

/**
 * Derive every name in the project from the web part title.
 *
 * Renaming these by hand afterwards is the tedious part of starting from a
 * sample: the folder, the class, the manifest alias, the strings specifier and
 * the bundle key all have to agree, and SharePoint only complains once at
 * runtime when they do not.
 */
export function toComponentNames(title: string): ComponentNames {
  const words = title
    .normalize("NFD")
    .replace(/[\u0300-\u036F]/g, "") // "Résumé" -> "Resume"
    .split(/[^A-Za-z0-9]+/)
    .filter(Boolean);

  let base = words.map((word) => word.charAt(0).toUpperCase() + word.slice(1)).join("");
  // Identifiers cannot start with a digit, and an empty title needs a fallback.
  if (!/^[A-Za-z]/.test(base)) base = `Wp${base}`;
  // "My Web Part" should not become "MyWebPartWebPart".
  base = base.replace(/WebPart$/, "") || base;

  const alias = `${base}WebPart`;
  return {
    dir: base.charAt(0).toLowerCase() + base.slice(1),
    component: RESERVED.has(base) ? `${base}Component` : base,
    alias,
    bundleName: alias.replace(/([a-z0-9])([A-Z])/g, "$1-$2").toLowerCase(),
    strings: `${alias}Strings`,
  };
}

/** "my-web-part" -> "My Web Part", for the title shown in SharePoint's toolbox. */
export function toTitle(input: string): string {
  const words = input
    .split(/[-_.\s]+/)
    .filter(Boolean)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1));
  return words.join(" ") || "Hello World";
}

/** A directory we may write into: missing, or present but empty. */
export function isUsable(dir: string): boolean {
  if (!fs.existsSync(dir)) return true;
  if (!fs.statSync(dir).isDirectory()) return false;
  return fs.readdirSync(dir).length === 0;
}
