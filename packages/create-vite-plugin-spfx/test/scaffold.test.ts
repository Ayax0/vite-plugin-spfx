import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";
import { afterAll, describe, expect, it } from "vitest";

import {
  isUsable,
  scaffold,
  substitute,
  toComponentNames,
  toPackageName,
  toTitle,
} from "../src/scaffold.ts";

const templatesDir = path.resolve(import.meta.dirname, "..", "templates");
const templates = fs.readdirSync(templatesDir);
const temp = fs.mkdtempSync(path.join(os.tmpdir(), "create-spfx-"));

afterAll(() => fs.rmSync(temp, { recursive: true, force: true }));

const names = toComponentNames("Team Roster");

const replacements = {
  __PROJECT_NAME__: "my-web-part",
  __TITLE__: "Team Roster",
  __FRAMEWORK__: "React",
  __COMPONENT_DIR__: names.dir,
  __COMPONENT__: names.component,
  __ALIAS__: names.alias,
  __BUNDLE_NAME__: names.bundleName,
  __STRINGS__: names.strings,
  __COMPONENT_ID__: randomUUID(),
  __SOLUTION_ID__: randomUUID(),
  __FEATURE_ID__: randomUUID(),
  __PLUGIN_VERSION__: "^0.1.0",
};

describe("substitute", () => {
  it("replaces known tokens", () => {
    expect(substitute("name: __PROJECT_NAME__", replacements)).toBe("name: my-web-part");
  });

  it("leaves unknown tokens alone rather than emptying them", () => {
    expect(substitute("__NOT_MINE__", replacements)).toBe("__NOT_MINE__");
  });

  it("does not touch Vue interpolation", () => {
    expect(substitute("{{ props.description }}", replacements)).toBe("{{ props.description }}");
  });
});

describe("naming", () => {
  it("makes an npm-safe package name", () => {
    expect(toPackageName("My Web Part")).toBe("my-web-part");
    // Separators are normalised to dashes, and a trailing one is dropped.
    expect(toPackageName("  Hello_World  ")).toBe("hello-world");
    expect(toPackageName("My Web Part!")).toBe("my-web-part");
    // Nothing usable left — fall back rather than emit an invalid name.
    expect(toPackageName("!!!")).toBe("spfx-web-part");
  });

  it("derives a readable title", () => {
    expect(toTitle("my-web-part")).toBe("My Web Part");
    expect(toTitle("hello")).toBe("Hello");
  });
});

describe("toComponentNames", () => {
  it("derives every name in the project from the title", () => {
    expect(toComponentNames("Team Roster")).toEqual({
      dir: "teamRoster",
      component: "TeamRoster",
      alias: "TeamRosterWebPart",
      bundleName: "team-roster-web-part",
      strings: "TeamRosterWebPartStrings",
    });
  });

  it("does not repeat 'web part' when the title already says it", () => {
    expect(toComponentNames("My Web Part").alias).toBe("MyWebPart");
    expect(toComponentNames("My Web Part").bundleName).toBe("my-web-part");
  });

  it("keeps the component from shadowing a framework import", () => {
    // `import * as React from "react"` plus a component called React is a
    // duplicate declaration, and the build fails before anything explains why.
    expect(toComponentNames("React").component).toBe("ReactComponent");
    expect(toComponentNames("Vue").component).toBe("VueComponent");
    // The class and folder are unaffected — only the component identifier clashes.
    expect(toComponentNames("React").alias).toBe("ReactWebPart");
    expect(toComponentNames("Reaction").component).toBe("Reaction");
  });

  it("survives punctuation, accents and a leading digit", () => {
    expect(toComponentNames("Résumé-Viewer!").component).toBe("ResumeViewer");
    // Identifiers cannot start with a digit.
    expect(toComponentNames("2026 Planner").component).toBe("Wp2026Planner");
    expect(toComponentNames("").alias).toBe("WpWebPart");
  });
});

describe("isUsable", () => {
  it("accepts a missing or empty directory and rejects a filled one", () => {
    const empty = path.join(temp, "empty");
    fs.mkdirSync(empty);
    expect(isUsable(path.join(temp, "missing"))).toBe(true);
    expect(isUsable(empty)).toBe(true);

    fs.writeFileSync(path.join(empty, "file.txt"), "");
    expect(isUsable(empty)).toBe(false);
  });
});

// The point of the suite: a scaffolded project must be a finished project, not
// a template with holes left in it.
describe.each(templates)("template %s", (template) => {
  const target = path.join(temp, template);
  const files = scaffold(path.join(templatesDir, template), target, replacements);

  const webpart = path.join("src", "webparts", names.dir);

  it("names the folder and the files after the web part, not after the sample", () => {
    expect(files).toEqual(
      expect.arrayContaining([
        "package.json",
        "tsconfig.json",
        "vite.config.mjs",
        path.join("config", "package-solution.json"),
        path.join("config", "serve.json"),
        path.join(webpart, `${names.alias}.ts`),
        path.join(webpart, `${names.alias}.manifest.json`),
      ]),
    );
    expect(files.some((file) => /HelloWorld/i.test(file))).toBe(false);
  });

  it("leaves no trace of the sample's names in any file", () => {
    const offenders = files.filter((file) =>
      /HelloWorld|hello-world|helloWorld/.test(fs.readFileSync(path.join(target, file), "utf8")),
    );
    expect(offenders).toEqual([]);
  });

  it("wires the plugin options to the generated names", () => {
    const config = fs.readFileSync(path.join(target, "vite.config.mjs"), "utf8");
    expect(config).toContain(`bundleName: "${names.bundleName}"`);
    expect(config).toContain(`entry: "src/webparts/${names.dir}/${names.alias}.ts"`);
    expect(config).toContain(`${names.strings}: "src/webparts/${names.dir}/loc/{locale}.js"`);
  });

  it("restores the .gitignore npm would have stripped", () => {
    expect(fs.existsSync(path.join(target, ".gitignore"))).toBe(true);
    expect(fs.existsSync(path.join(target, "_gitignore"))).toBe(false);
  });

  it("leaves no placeholder behind", () => {
    const leftovers = files.filter((file) =>
      /__[A-Z0-9_]+__/.test(fs.readFileSync(path.join(target, file), "utf8")),
    );
    expect(leftovers).toEqual([]);
  });

  it("writes a package.json that depends on the plugin", () => {
    const pkg = JSON.parse(fs.readFileSync(path.join(target, "package.json"), "utf8"));
    expect(pkg.name).toBe("my-web-part");
    expect(pkg.private).toBe(true);
    expect(pkg.devDependencies["vite-plugin-spfx"]).toBe("^0.1.0");
    expect(pkg.scripts).toMatchObject({ dev: "vite", build: "vite build" });
  });

  it("gives the component the generated id, not the example's", () => {
    const manifest = JSON.parse(
      fs.readFileSync(path.join(target, webpart, `${names.alias}.manifest.json`), "utf8"),
    );
    expect(manifest.id).toBe(replacements["__COMPONENT_ID__"]);
    expect(manifest.alias).toBe(names.alias);
    expect(manifest.preconfiguredEntries[0].title.default).toBe("Team Roster");

    const solution = JSON.parse(
      fs.readFileSync(path.join(target, "config", "package-solution.json"), "utf8"),
    );
    expect(solution.solution.id).toBe(replacements["__SOLUTION_ID__"]);
    expect(solution.solution.id).not.toBe(manifest.id);
    expect(solution.paths.zippedPackage).toBe("solution/my-web-part.sppkg");
  });
});
