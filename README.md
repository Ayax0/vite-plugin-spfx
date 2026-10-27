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

```sh
npm create vite-plugin-spfx@latest
```

Pick a framework, and you have a project that builds a deployable `.sppkg`.
Or add the plugin to a Vite project you already have:

```sh
npm install -D vite-plugin-spfx
```

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

`vite build` emits the AMD bundle, the component manifest and a deployable
`.sppkg`. `vite dev` serves them into SharePoint's debug flow.

**→ [Full documentation](./packages/vite-plugin-spfx#readme)** — options, theming,
localization, dev server, packaging.

## Templates

`npm create` scaffolds the same web part five ways, built by the same plugin
without a line of difference in its code.

| | Bundle | Framework comes from | Styles |
| --- | --- | --- | --- |
| [react](./packages/create-vite-plugin-spfx/templates/react/) | **2.6 kB** | SharePoint | `.module.scss` |
| [vanilla](./packages/create-vite-plugin-spfx/templates/vanilla/) | 2.9 kB | — no framework | `.module.scss` |
| [solid](./packages/create-vite-plugin-spfx/templates/solid/) | 17.4 kB | the bundle | `.module.scss` |
| [svelte](./packages/create-vite-plugin-spfx/templates/svelte/) | 37.6 kB | the bundle | `<style>` |
| [vue](./packages/create-vite-plugin-spfx/templates/vue/) | 62.7 kB | the bundle | `<style scoped>` |

React is tiny because SharePoint already ships it — the plugin discovers that and
references it instead of bundling it. Vanilla shows the floor: SharePoint only
asks for a class that renders into a DOM element, and no framework at all
satisfies that.

Each one covers the property pane, localization in two locales, a
runtime-resolved asset, theming, and component state that survives the re-render
SPFx triggers on every property pane keystroke. They differ only in the component
file and in how props are updated.

The template sources carry `__PLACEHOLDER__` tokens, so read them as a shape
rather than as a project — scaffold one to see the real thing.

## Repository

```
packages/vite-plugin-spfx/          the plugin — published to npm
packages/create-vite-plugin-spfx/   the scaffolder behind `npm create`
  templates/                        the five projects it writes
```

```sh
pnpm install
pnpm build
pnpm test
pnpm --filter create-vite-plugin-spfx verify-templates   # scaffold, install, build all five
```

Requires Node ≥ 22.12 and pnpm. `verify-templates` is what proves the templates
still produce a working project: it drives the real CLI, installs the plugin from
a packed tarball the way a user would, and fails unless a bundle and a `.sppkg`
come out. CI runs it on every push.

## Docs

- [CONTRIBUTING.md](./CONTRIBUTING.md) — source layout, tests, release
- [Plugin README](./packages/vite-plugin-spfx#readme) — options, theming,
  localization, dev server, packaging
- [Scaffolder README](./packages/create-vite-plugin-spfx#readme) — `npm create`
  options and what each template contains

## Status

✅ shipped · 🧭 specified, not built · ⛔ not planned

### Component types

| | | |
| :-: | --- | --- |
| ✅ | **Web parts** | Property pane, theming, localization, assets |
| 🧭 | Extensions | Application customizer, field customizer, command set — element XML and query params are understood |
| 🧭 | Adaptive card extensions | |
| 🧭 | Library components | |

### Build

| | | |
| :-: | --- | --- |
| ✅ | **Named AMD bundle** | Synthesised from Vite's ESM; neither Vite 8 nor Rolldown can emit AMD |
| ✅ | **Component manifest** | `loaderConfig` and `scriptResources`, dev and release |
| ✅ | **Component discovery** | Checks every import against the installed package, so SPFx version bumps need no plugin update |
| ✅ | **`.sppkg` packaging** | Byte-reproducible from `config/package-solution.json` |
| ✅ | **Localization** | One AMD module per locale, `localizedPath` in the manifest |
| ✅ | **Theming** | `[theme:…]` tokens as CSS custom properties, no `load-themed-styles` |
| ✅ | **Assets** | Resolved at runtime from the script's own URL |
| 🧭 | Multi-component bundles | One component per bundle today; the `exportName` form is specified but unbuilt |
| 🧭 | Code splitting | Currently `inlineDynamicImports`; real chunk loading needs a CSP test on SharePoint Online first |

### Dev loop

| | | |
| :-: | --- | --- |
| ✅ | **HTTPS dev server** | Debug protocol for the hosted workbench and real site pages |
| ✅ | **Live reload** | ~100 ms rebuild, then the page reloads itself |
| ✅ | **Certificate** | Reuses the Rushstack store, so a machine set up for Heft needs nothing |
| ⛔ | HMR | SharePoint's loader is a RequireJS fork with no ESM pathway. `heft start` is a full reload too |

### Verified

Against SharePoint Online, SPFx 1.23.2 and Vite 8: hosted workbench and real
site pages, app catalog install, theming, localization, the property pane, and
a `.sppkg` built clean-room that the app catalog accepts. The manifest,
`scriptResources` and package inventory are checked byte-for-byte against
Microsoft's own toolchain output in [the test suite](./packages/vite-plugin-spfx/test/).

## License

[MIT](./LICENSE)
