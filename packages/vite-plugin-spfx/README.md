<div align="center">

# vite-plugin-spfx

### Ship SharePoint web parts at Vite speed.

One plugin instead of the SPFx toolchain —<br>
in React, Vue, Svelte or whatever you reach for next.<br>
**No Heft. No webpack. No Gulp. ~1050 fewer packages.**

[![npm version](https://img.shields.io/npm/v/vite-plugin-spfx?color=yellow&style=flat-square)](https://npmjs.com/package/vite-plugin-spfx)
[![npm downloads](https://img.shields.io/npm/dm/vite-plugin-spfx?color=yellow&style=flat-square)](https://npmjs.com/package/vite-plugin-spfx)
[![license](https://img.shields.io/npm/l/vite-plugin-spfx?color=yellow&style=flat-square)](./LICENSE)

</div>

---

A guide from an empty folder to a web part running on a SharePoint page. Already
have an SPFx project? Skip to
[Migrating an existing project](#migrating-an-existing-project).

> [!TIP]
> To skip the setup entirely, scaffold it:
> `npm create vite-plugin-spfx@latest`. You get the finished project this guide
> builds by hand, in your choice of framework. The rest of this page then works
> as a reference for what the generated files are for.

**Contents** — [Requirements](#requirements) · [1. Set up the project](#1-set-up-the-project) ·
[2. Configure Vite](#2-configure-vite) · [3. Write the web part](#3-write-the-web-part) ·
[4. Run it against SharePoint](#4-run-it-against-sharepoint) · [5. Ship it](#5-ship-it) ·
[Styles and theming](#styles-and-theming) · [Localization](#localization) ·
[Assets](#assets) · [Framework notes](#framework-notes) · [Options](#options) ·
[Troubleshooting](#troubleshooting)

## Requirements

|            |                                                                                                                                                                |
| ---------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Node       | ≥ 22.12                                                                                                                                                        |
| Vite       | 7 or 8                                                                                                                                                         |
| SPFx       | Verified against 1.23.2. Component discovery reads the installed packages rather than a hardcoded table, so other versions are expected to work — but untested |
| SharePoint | SharePoint Online                                                                                                                                              |
| Access     | Ability to upload a `.sppkg` to an app catalog, for anything beyond the workbench                                                                              |

You do not need a tenant to build — only to run what you built.

## 1. Set up the project

An SPFx component is four things: a **web part class**, a **component manifest**
describing it to SharePoint, a **solution config** describing the package, and a
**serve config** for the dev loop. The plugin reads all four in their stock SPFx
format, so nothing below is a plugin-specific invention.

```
my-web-part/
├─ config/
│  ├─ package-solution.json      what the .sppkg contains
│  └─ serve.json                 port and workbench URL for the dev loop
├─ src/webparts/helloWorld/
│  ├─ HelloWorldWebPart.ts       the component
│  └─ HelloWorldWebPart.manifest.json
├─ package.json
├─ tsconfig.json
└─ vite.config.mjs
```

### Dependencies

```sh
npm install -D vite vite-plugin-spfx typescript
npm install @microsoft/sp-core-library @microsoft/sp-component-base \
            @microsoft/sp-property-pane @microsoft/sp-webpart-base
```

Pin the `@microsoft/sp-*` packages to the SPFx version your tenant runs, and keep
them on the same version as each other.

> [!IMPORTANT]
> These packages must be real, installed dependencies even though they end up
> _outside_ your bundle. The plugin discovers them by reading the
> `dist/<guid>.manifest.json` that each one ships, and that is where the
> component id and version in your output come from. Nothing is hardcoded, which
> is why an SPFx version bump needs no plugin update.

Add `sass` if you write `.scss`, plus your framework's Vite plugin if you use one
— with [one exception for React](#framework-notes).

### `package.json`

```json
{
  "name": "my-web-part",
  "version": "1.0.0",
  "private": true,
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "vite build"
  }
}
```

`version` matters: it fills the `"*"` in your component manifest and becomes part
of the AMD module id SharePoint resolves.

### `tsconfig.json`

```json
{
  "compilerOptions": {
    "target": "ES2017",
    "lib": ["DOM", "ES2017"],
    "module": "ESNext",
    "moduleResolution": "bundler",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "noEmit": true,
    "types": ["vite/client"]
  },
  "include": ["src", "vite.config.mjs"]
}
```

`noEmit` is deliberate — Vite compiles, `tsc` only type-checks. There is no
two-stage `lib/` build the way Heft does it.

### `config/package-solution.json`

The stock SPFx schema. Generate fresh GUIDs for `solution.id` and every
`features[].id`.

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
  "solution": {
    "name": "my-web-part",
    "id": "b7e78259-b6bf-415e-99b1-411758dcb280",
    "version": "1.0.0.0",
    "includeClientSideAssets": true,
    "skipFeatureDeployment": true,
    "isDomainIsolated": false,
    "features": [
      {
        "title": "My web part",
        "description": "Deploys the hello world web part.",
        "id": "4170584d-59e3-47c2-b953-02d022dd4740",
        "version": "1.0.0.0"
      }
    ]
  },
  "paths": {
    "zippedPackage": "solution/my-web-part.sppkg"
  }
}
```

`skipFeatureDeployment: true` makes the web part available tenant-wide on
install, instead of requiring per-site activation. Without this file the build
still works — you simply get no `.sppkg`.

### `config/serve.json`

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json",
  "port": 4321,
  "https": true,
  "initialPage": "https://{tenantDomain}/_layouts/workbench.aspx"
}
```

`{tenantDomain}` stays a placeholder on purpose — it is filled from your
environment, so this file can be committed. The defaults are port `4321` and
HTTPS on, which makes the file optional.

### The component manifest

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json",
  "id": "1c55fdd9-ba6d-43ef-ba3d-fa7a5d920dcc",
  "alias": "HelloWorldWebPart",
  "componentType": "WebPart",
  "version": "*",
  "manifestVersion": 2,
  "requiresCustomScript": false,
  "supportedHosts": ["SharePointWebPart"],
  "supportsThemeVariants": true,
  "preconfiguredEntries": [
    {
      "groupId": "5c03119e-3074-46fd-976b-c60198311f70",
      "group": { "default": "Advanced" },
      "title": { "default": "Hello World" },
      "description": { "default": "My first Vite-built web part" },
      "officeFabricIconFontName": "Page",
      "properties": { "description": "Hello" }
    }
  ]
}
```

`id` is a GUID you generate once and never change — SharePoint identifies the
component by it, and changing it orphans every placed instance.
`preconfiguredEntries` is what appears in the toolbox.

## 2. Configure Vite

```js
// vite.config.mjs
import { defineConfig } from "vite";
import { spfx } from "vite-plugin-spfx";

export default defineConfig({
  plugins: [
    spfx({
      bundleName: "hello-world-web-part",
      entry: "src/webparts/helloWorld/HelloWorldWebPart.ts",
      manifest: "src/webparts/helloWorld/HelloWorldWebPart.manifest.json",
    }),
  ],
});
```

Three required options and you are done. `bundleName` is the key SPFx calls the
bundle by — it becomes the output file name and `loaderConfig.entryModuleId`.

Everything the SharePoint loader depends on — module format, externals, entry
handling, asset URLs — is configured by the plugin. See
[what the plugin configures](#what-the-plugin-configures-for-you) for the exact
list and which parts you can still override.

## 3. Write the web part

Default-export a class extending `BaseClientSideWebPart` that renders into
`this.domElement`. That is SharePoint's entire contract — which is why any
framework works.

```ts
// src/webparts/helloWorld/HelloWorldWebPart.ts
import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import { Version } from "@microsoft/sp-core-library";
import {
  PropertyPaneTextField,
  type IPropertyPaneConfiguration,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { applyTheme } from "vite-plugin-spfx/runtime";

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  public render(): void {
    this.domElement.textContent = `Hello, ${this.context.pageContext.user.displayName}`;
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    applyTheme(this.domElement, theme);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [PropertyPaneTextField("description", { label: "Description" })],
            },
          ],
        },
      ],
    };
  }
}
```

> [!WARNING]
> **SPFx calls `render()` on every property pane change** — on every keystroke in
> a text field. If `render()` mounts a fresh component tree each time, component
> state resets while the user types. React's `ReactDom.render()` reconciles into
> the same container and hides this; everywhere else, mount once in the first
> `render()` and update reactive state on later calls.

## 4. Run it against SharePoint

```sh
npx vite
```

This starts an HTTPS server and a watch build together, then prints the link to
open:

```
  ➜  Component: Hello World
  ➜  Open:      https://contoso.sharepoint.com/_layouts/workbench.aspx?debugManifestsFile=…
  ➜  Query:     ?debugManifestsFile=…&debug=true&noredir=true
```

There is no HMR: SharePoint's loader has no ESM pathway, so a rebuild triggers a
page reload — about 100 ms, then the page reloads itself. `heft start`
effectively does the same.

### Your tenant

The tenant is per-developer data, so it lives outside committed config. Put it in
`.env.local`:

```sh
SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com
```

Then `vite --open` lands on the workbench directly. Without it you still get the
query string to append to a SharePoint page yourself.

### The certificate

> [!IMPORTANT]
> The certificate must be **trusted**, not merely present. Your bundle is loaded
> cross-origin as a subresource from an https page: browsers reject an untrusted
> certificate outright and offer no click-through, and plain http is blocked as
> mixed content. A self-signed throwaway will not do.

The plugin uses the same certificate store the SPFx tooling does, so a machine
already set up for Heft needs nothing. Otherwise:

```sh
npx spfx-trust-dev-cert
```

Installing a local certificate authority is a one-off privileged step, so its
dependencies are optional rather than something every install and CI run carries:

```sh
npm install -D @rushstack/debug-certificate-manager @rushstack/terminal
```

Bringing your own works too — `server.https` in `vite.config` always wins.

### Debugging on a real page

The hosted workbench is being retired, but the same debug protocol works on any
modern SharePoint page: append the printed `Query` string to the page URL.
Chromium may additionally ask you to allow local network access.

## 5. Ship it

```sh
npx vite build
```

| Output                        | What it is                                                        |
| ----------------------------- | ----------------------------------------------------------------- |
| `dist/`                       | The AMD bundle, the component manifest, locale modules and assets |
| `release/`                    | The same files, arranged the way the packager expects             |
| `sharepoint/solution/*.sppkg` | The deployable package                                            |

Upload the `.sppkg` to your tenant app catalog, approve it, then add the web part
from the toolbox. With `includeClientSideAssets: true` the assets travel inside
the package, so no CDN is involved — set [`cdnBasePath`](#options) if you host
them elsewhere.

The package is byte-reproducible: identical input produces an identical file,
because timestamps are fixed and the internal GUIDs are derived from your
solution id. Heft randomises both on every build, which makes its output
impossible to diff.

> [!NOTE]
> A `.sppkg` is only written for production builds — `vite build` is production
> by default. A development build stops after `dist/` and `release/`, the same
> way Heft refuses to package a debug build.

## Styles and theming

Write styles however Vite lets you: `.scss`, CSS modules, a framework's scoped
style block. The plugin inlines the finished CSS into the bundle and injects it
at runtime, because SPFx manifests have no CSS resource type. No stylesheet is
emitted, and none is expected.

For theme-aware colours, use the SPFx token syntax:

```scss
.hello {
  color: "[theme:bodyText, default: #323130]";
}
```

That is rewritten to `var(--bodyText, #323130)` at build time. Publish the host's
theme slots as those custom properties:

```ts
import { applyTheme } from "vite-plugin-spfx/runtime";

protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
  applyTheme(this.domElement, theme);
}
```

Stock SPFx expects one `style.setProperty()` per slot inside `onThemeChanged` —
easy to forget, and silently outdated as soon as your CSS uses a new slot.
`applyTheme` publishes every slot the host provides, so the CSS decides which
ones matter. The `@microsoft/load-themed-styles` runtime is not involved.

> [!NOTE]
> **Migrating:** in SPFx, _every_ `.scss` file is a CSS module unless named
> `.global.scss`. Under Vite the usual convention applies — only `*.module.scss`
> is a module. Rename accordingly, or your class names stop being hashed.

## Localization

```js
spfx({
  localizedResources: {
    HelloWorldWebPartStrings: "src/webparts/helloWorld/loc/{locale}.js",
  },
});
```

Every file matching the pattern ships as a standalone AMD module. Locale files
stay exactly what SPFx uses — hand-written AMD, copied verbatim:

```js
// src/webparts/helloWorld/loc/en-us.js
define([], function () {
  return { PropertyPaneDescription: "Settings for this web part" };
});
```

The type comes from a hand-maintained declaration next to them. Nothing is
generated:

```ts
// src/webparts/helloWorld/loc/mystrings.d.ts
declare interface IHelloWorldWebPartStrings {
  PropertyPaneDescription: string;
}

declare module "HelloWorldWebPartStrings" {
  const strings: IHelloWorldWebPartStrings;
  export = strings;
}
```

Then `import * as strings from "HelloWorldWebPartStrings"` resolves through the
manifest at runtime.

With more than one locale the manifest gets a `localizedPath` entry and
SharePoint picks by page language, falling back to `en-US`. Manifest keys are
normalised with SPFx's own algorithm (`de-de` → `de-DE`, including script subtags
and pseudo-locales) because SharePoint matches the page language against those
keys — the casing is not cosmetic.

## Assets

Import them and Vite handles the rest:

```ts
import logoUrl from "./assets/logo.svg";
```

At runtime the URL resolves against the script's own location, so the same bundle
works from `localhost`, from the app catalog and from a CDN without rebuilding.
Assets under the inline limit become data URIs as usual; set
`build.assetsInlineLimit: 0` to keep them as real files.

## Framework notes

> [!WARNING]
> **React** — do not use `@vitejs/plugin-react`. Its automatic JSX runtime
> imports `react/jsx-runtime`, a specifier SharePoint does not serve. Use the
> classic runtime instead: `"jsx": "react"` in `tsconfig.json`, and import React
> yourself. React comes from SharePoint rather than your bundle, which is why a
> React web part is the smallest of the three.

**Everything else** — Vue, Svelte, Solid and friends work through their normal
Vite plugin and are bundled, since SharePoint provides only React.

**No framework** — also a legitimate choice. SharePoint asks for a class that
renders into a DOM element, nothing more, and skipping the framework produces the
smallest bundle after React.

**All of them** — mind the `render()`-on-every-keystroke behaviour described in
[step 3](#3-write-the-web-part). `npm create vite-plugin-spfx@latest` scaffolds
the same web part in React, Vue, Svelte, Solid or plain TypeScript, each already
handling that correctly for its framework — the quickest way to see the pattern
is to generate one.

## Migrating an existing project

Your `config/package-solution.json`, `config/serve.json` and component manifests
work unchanged — the plugin reads the stock formats. The component model,
`BaseClientSideWebPart`, the property pane and the `.sppkg` your app catalog
expects all stay the same. What you drop is the build toolchain.

What to change:

1. Delete `config/config.json`, `config/deploy-azure-storage.json`, the
   toolchain devDependencies (`@microsoft/spfx-*`, `@rushstack/heft`, or the
   older `gulpfile.js` and its plugins), and add `vite` plus `vite-plugin-spfx`.
2. Add `vite.config.mjs` with one `spfx()` entry per bundle from your old
   `config.json` — one component per bundle is the current limit.
3. Set `"noEmit": true` in `tsconfig.json` and drop the SPFx rig `extends`.
4. Rename `.scss` files to `.module.scss` wherever you rely on scoped class names
   (see [Styles and theming](#styles-and-theming)).
5. Replace the per-slot `setProperty` calls in `onThemeChanged` with `applyTheme`.
6. Move your tenant out of `serve.json` into `.env.local`.

Locale files, component manifests, GUIDs and the solution config need no edits.

## Options

| Option               | Type                     | Description                                                                                                             |
| -------------------- | ------------------------ | ----------------------------------------------------------------------------------------------------------------------- |
| `bundleName`         | `string`                 | **Required.** Bundle key, as in SPFx's `config.json`. Becomes `loaderConfig.entryModuleId` and the output file name.    |
| `entry`              | `string`                 | **Required.** Component entry, relative to the project root.                                                            |
| `manifest`           | `string`                 | **Required.** The component's `*.manifest.json`.                                                                        |
| `localizedResources` | `Record<string, string>` | Bare specifier → per-locale files, e.g. `{ MyStrings: "src/.../loc/{locale}.js" }`.                                     |
| `serveBaseUrl`       | `string`                 | Base URL written into the debug manifest. Defaults to the dev server's own origin.                                      |
| `cdnBasePath`        | `string`                 | Base URL written into the release manifest. Empty ships assets inside the package.                                      |
| `serveInitialPage`   | `string`                 | Page the dev server links to, e.g. `https://{tenantDomain}/_layouts/workbench.aspx`. Falls back to `config/serve.json`. |

### What the plugin configures for you

**Fixed** — these follow from the SharePoint loader, and changing them breaks the
output: `build.cssCodeSplit: false`, `build.modulePreload: false`, the entry, the
externals, `preserveEntrySignatures: "strict"`, ESM output with code splitting
off, the output file naming, and a PostCSS pass for theme tokens.

**Defaults** — chosen for SharePoint, yours to override:

| Option            | Default            | Why                                                  |
| ----------------- | ------------------ | ---------------------------------------------------- |
| `build.target`    | `es2017`           | SharePoint Online serves that level itself           |
| `build.minify`    | on for production  | readable while developing, minified to ship          |
| `build.sourcemap` | on for development | mirrors Heft, which emits no maps for release builds |

Vite merges a plugin's config _over_ yours, so a plugin that returned these
unconditionally would make them unchangeable. This one reads your config first
and only fills the gaps.

## Troubleshooting

**`external "x" has no scriptResources entry`** — you imported a bare specifier
the SharePoint loader cannot resolve. Either install it as a dependency, if
SharePoint serves it, or let it be bundled instead of externalised.

**The web part stays blank and the console shows a script error** — almost always
the certificate. Open `https://localhost:4321/` directly; if the browser
complains there, SharePoint never had a chance. See
[the certificate](#the-certificate).

**A property pane field loses focus or resets after one character** — `render()`
is remounting. See the warning in [step 3](#3-write-the-web-part).

**Styles disappear after deploying but work in dev** — theme tokens are rewritten
at build time and the CSS is inlined into the bundle. A CSS post-processor of
your own has to run before the theme pass, not after.

**`no locale files matched …`** — the `{locale}` pattern found nothing. The path
is relative to the project root, and the placeholder matches one whole file name
segment.

## Not supported yet

Web parts only, one component per bundle, no code splitting. Extensions, adaptive
card extensions, library components and multi-component bundles are understood
but unimplemented.

## License

[MIT](./LICENSE)
