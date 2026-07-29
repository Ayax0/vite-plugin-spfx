#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";
import { parseArgs, styleText } from "node:util";
import { randomUUID } from "node:crypto";

import { PromptCancelledError, select, text } from "./prompts.ts";
import { isUsable, scaffold, toComponentNames, toPackageName, toTitle } from "./scaffold.ts";

const TEMPLATES = [
  { value: "react", label: "React", hint: "smallest — SharePoint serves React itself" },
  { value: "vanilla", label: "None", hint: "plain TypeScript, no framework" },
  { value: "solid", label: "Solid", hint: "" },
  { value: "svelte", label: "Svelte", hint: "" },
  { value: "vue", label: "Vue", hint: "" },
] as const;

type Template = (typeof TEMPLATES)[number]["value"];

const root = path.resolve(import.meta.dirname, "..");
const { version } = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8")) as {
  version: string;
};

const bold = (value: string): string => styleText("bold", value);
const cyan = (value: string): string => styleText("cyan", value);
const dim = (value: string): string => styleText("dim", value);

const HELP = `
  ${bold("create-vite-plugin-spfx")} ${dim(`v${version}`)}

  Scaffold a SharePoint Framework web part built by Vite.

  ${bold("Usage")}
    npm create vite-plugin-spfx@latest [directory] [options]

  ${bold("Options")}
    -t, --template <name>   ${TEMPLATES.map((entry) => entry.value).join(", ")}
    -y, --yes               accept the defaults, ask nothing
    -h, --help              show this message
    -v, --version           print the version
`;

async function main(): Promise<void> {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      template: { type: "string", short: "t" },
      yes: { type: "boolean", short: "y", default: false },
      help: { type: "boolean", short: "h", default: false },
      version: { type: "boolean", short: "v", default: false },
    },
  });

  if (values.help) return void console.log(HELP);
  if (values.version) return void console.log(version);

  if (values.template && !TEMPLATES.some((entry) => entry.value === values.template)) {
    throw new Error(
      `unknown template "${values.template}" — expected one of ${TEMPLATES.map((entry) => entry.value).join(", ")}`,
    );
  }

  const defaultDir = positionals[0] ?? "my-web-part";
  const directory =
    positionals[0] ?? (values.yes ? defaultDir : await text("Project directory", defaultDir));

  const template = (values.template ??
    (values.yes ? "react" : await select("Framework", [...TEMPLATES]))) as Template;

  const targetDir = path.resolve(process.cwd(), directory);
  if (!isUsable(targetDir)) {
    throw new Error(`${path.relative(process.cwd(), targetDir) || "."} exists and is not empty`);
  }

  const name = toPackageName(path.basename(targetDir));
  const title = values.yes ? toTitle(name) : await text("Web part title", toTitle(name));

  const names = toComponentNames(title);

  // Fresh ids on every scaffold. Reusing them across projects would make
  // SharePoint treat two different web parts as the same component.
  const files = scaffold(path.join(root, "templates", template), targetDir, {
    __PROJECT_NAME__: name,
    __TITLE__: title,
    __FRAMEWORK__: TEMPLATES.find((entry) => entry.value === template)!.label,
    __COMPONENT_DIR__: names.dir,
    __COMPONENT__: names.component,
    __ALIAS__: names.alias,
    __BUNDLE_NAME__: names.bundleName,
    __STRINGS__: names.strings,
    __COMPONENT_ID__: randomUUID(),
    __SOLUTION_ID__: randomUUID(),
    __FEATURE_ID__: randomUUID(),
    __PLUGIN_VERSION__: `^${version}`,
  });

  const relative = path.relative(process.cwd(), targetDir);
  const runner = packageManager();

  console.log(`
  ${bold(`Created ${title}`)} ${dim(`(${files.length} files)`)}

  ${bold("Next")}
    ${relative ? `cd ${cyan(relative)}` : ""}
    ${cyan(`${runner} install`)}
    ${cyan(`${runner} run dev`)}      ${dim("serve it to SharePoint")}
    ${cyan(`${runner} run build`)}    ${dim("build dist/ and the .sppkg")}

  ${dim("The dev server needs a trusted certificate and your tenant domain —")}
  ${dim("the generated README walks through both.")}
`);
}

/** Whichever tool invoked us, so the printed commands are the ones that will work. */
function packageManager(): string {
  const agent = process.env["npm_config_user_agent"] ?? "";
  if (agent.startsWith("pnpm")) return "pnpm";
  if (agent.startsWith("yarn")) return "yarn";
  if (agent.startsWith("bun")) return "bun";
  return "npm";
}

try {
  await main();
} catch (error) {
  if (error instanceof PromptCancelledError) {
    console.log(dim("\n  cancelled"));
    process.exit(130);
  }
  console.error(
    styleText("red", `\n  ${error instanceof Error ? error.message : String(error)}\n`),
  );
  process.exit(1);
}
