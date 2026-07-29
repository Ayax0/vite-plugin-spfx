<div align="center">

# create-vite-plugin-spfx

### One command to a SharePoint web part.

Scaffolds a SharePoint Framework project built by Vite —<br>
React, Vue, Svelte, Solid, or no framework at all.

[![npm version](https://img.shields.io/npm/v/create-vite-plugin-spfx?color=yellow&style=flat-square)](https://npmjs.com/package/create-vite-plugin-spfx)
[![license](https://img.shields.io/npm/l/create-vite-plugin-spfx?color=yellow&style=flat-square)](./LICENSE)

</div>

---

```sh
npm create vite-plugin-spfx@latest
# or
pnpm create vite-plugin-spfx
```

It asks three things — directory, framework, web part title — and writes a
project that builds a deployable `.sppkg` right away:

```sh
cd my-web-part
npm install
npm run dev     # serve it to SharePoint
npm run build   # dist/ and sharepoint/solution/my-web-part.sppkg
```

## Non-interactive

```sh
npm create vite-plugin-spfx@latest my-web-part -- --template solid --yes
```

| Option                  |                                                                      |
| ----------------------- | -------------------------------------------------------------------- |
| `[directory]`           | Where to write. Must be empty or missing. Defaults to `my-web-part`. |
| `-t, --template <name>` | `react`, `vanilla`, `solid`, `svelte`, `vue`                         |
| `-y, --yes`             | Take every default, ask nothing                                      |
| `-h, --help`            | Show usage                                                           |
| `-v, --version`         | Print the version                                                    |

## Which template?

|             |                                                                                           |
| ----------- | ----------------------------------------------------------------------------------------- |
| **react**   | Smallest bundle — SharePoint serves React itself, so it is referenced rather than bundled |
| **vanilla** | No framework at all. SharePoint only asks for a class that renders into a DOM element     |
| **solid**   | Fine-grained reactivity, compiles JSX to direct DOM calls                                 |
| **svelte**  | Compiles away, contributes a small runtime                                                |
| **vue**     | Full renderer, largest of the five                                                        |

All five build the same web part — property pane, localization in two locales,
a runtime-resolved asset, theming, and component state that survives a
re-render. They differ only in the component file and in how props are updated.

## What you get

Everything is named after your web part, so there is nothing to rename before
you start. For the title _Team Roster_:

```
team-roster/
├─ config/
│  ├─ package-solution.json      what the .sppkg contains
│  └─ serve.json                 port and workbench URL
├─ src/webparts/teamRoster/
│  ├─ TeamRosterWebPart.ts       the web part class
│  ├─ TeamRosterWebPart.manifest.json
│  ├─ TeamRoster.vue             the component
│  ├─ loc/                       en-us and de-de
│  └─ assets/
├─ package.json
├─ tsconfig.json
└─ vite.config.mjs
```

| Derived from the title   |                            |
| ------------------------ | -------------------------- |
| Folder                   | `src/webparts/teamRoster/` |
| Class and manifest alias | `TeamRosterWebPart`        |
| Component                | `TeamRoster`               |
| Strings specifier        | `TeamRosterWebPartStrings` |
| Bundle                   | `team-roster-web-part.js`  |

Those five have to agree with each other, and SharePoint only complains at
runtime when they do not — which is exactly why the scaffolder does it rather
than leaving you a sample to rename.

Fresh GUIDs are generated for the component, the solution and the feature on
every scaffold — reusing them would make SharePoint treat two different web
parts as the same component.

The generated README covers the two things the dev server needs from you: your
tenant domain in `.env.local`, and a trusted development certificate.

## Requirements

Node ≥ 22.12. A SharePoint tenant is only needed to run what you built, not to
build it.

## Documentation

The plugin doing the actual work is
[vite-plugin-spfx](https://github.com/Ayax0/vite-plugin-spfx/tree/main/packages/vite-plugin-spfx#readme)
— options, theming, localization, dev server, packaging, and how to migrate an
existing SPFx project.

## License

[MIT](./LICENSE)
