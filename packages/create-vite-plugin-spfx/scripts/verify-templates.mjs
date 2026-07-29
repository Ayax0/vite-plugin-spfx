#!/usr/bin/env node
/**
 * Scaffold every template, install it, and build it.
 *
 * The templates are the only copy of these projects, so nothing else proves
 * they still work. This drives the real CLI, installs the plugin from a packed
 * tarball rather than a workspace link — which is what a user gets from the
 * registry — and fails unless a bundle and a `.sppkg` come out the other end.
 *
 * Slow by nature (a full install per template), so it runs in CI and on demand
 * rather than as part of `pnpm test`.
 */
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const here = path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, "$1"));
const packageRoot = path.resolve(here, "..");
const pluginDir = path.resolve(packageRoot, "..", "vite-plugin-spfx");
const cli = path.join(packageRoot, "dist", "index.js");
const templatesDir = path.join(packageRoot, "templates");

const npm = process.platform === "win32" ? "npm.cmd" : "npm";
const only = process.argv[2];

if (!fs.existsSync(cli)) {
  console.error(`missing ${path.relative(packageRoot, cli)} — run: pnpm build`);
  process.exit(1);
}

const workDir = fs.mkdtempSync(path.join(os.tmpdir(), "verify-templates-"));
const templates = fs.readdirSync(templatesDir).filter((name) => !only || name === only);
const failures = [];

console.log(`packing vite-plugin-spfx`);
run(npm, ["pack", "--pack-destination", workDir], pluginDir);
const tarball = path.join(
  workDir,
  fs.readdirSync(workDir).find((file) => file.endsWith(".tgz")),
);

for (const template of templates) {
  // Deliberately not named after the template: a project called "react" would
  // derive a component called React and collide with the framework import — a
  // real edge case, but one the unit tests cover. Here we want the normal path.
  const projectDir = path.join(workDir, `team-roster-${template}`);
  console.log(`\n=== ${template} ===`);

  try {
    run(process.execPath, [cli, projectDir, "--template", template, "--yes"], workDir);

    // Point the dependency at the tarball: the version the templates pin is not
    // on the registry until release, and a workspace link would not prove the
    // published package works.
    const manifest = path.join(projectDir, "package.json");
    const pkg = JSON.parse(fs.readFileSync(manifest, "utf8"));
    pkg.devDependencies["vite-plugin-spfx"] = `file:${tarball.replaceAll("\\", "/")}`;
    fs.writeFileSync(manifest, `${JSON.stringify(pkg, undefined, 2)}\n`);

    run(npm, ["install", "--no-audit", "--no-fund"], projectDir);
    run(npm, ["run", "build"], projectDir);

    const bundles = fs
      .readdirSync(path.join(projectDir, "dist"))
      .filter((file) => file.endsWith(".js"));
    const packages = fs.readdirSync(path.join(projectDir, "sharepoint", "solution"));
    if (bundles.length === 0) throw new Error("no bundle in dist/");
    if (!packages.some((file) => file.endsWith(".sppkg"))) throw new Error("no .sppkg");

    console.log(`ok — ${bundles.length} bundle(s), ${packages.join(", ")}`);
  } catch (error) {
    failures.push(`${template}: ${error.message}`);
    console.error(`FAILED — ${error.message}`);
  }
}

if (failures.length > 0) {
  console.error(`\n${failures.length} of ${templates.length} templates failed:`);
  for (const failure of failures) console.error(`  ${failure}`);
  console.error(`\nkept for inspection: ${workDir}`);
  process.exit(1);
}

fs.rmSync(workDir, { recursive: true, force: true });
console.log(`\nall ${templates.length} templates scaffold, install and build`);

/**
 * Node refuses to spawn a `.cmd` without a shell (CVE-2024-27980), which is what
 * `npm` is on Windows. Shelling out means the arguments are re-parsed by cmd, so
 * anything with a space has to be quoted — temp paths regularly have one.
 */
function run(command, args, cwd) {
  const shell = command.endsWith(".cmd");
  const quoted = shell ? args.map((arg) => (/[\s&|<>^"]/.test(arg) ? `"${arg}"` : arg)) : args;
  execFileSync(command, quoted, { cwd, stdio: "inherit", shell });
}
