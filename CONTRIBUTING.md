# Contributing

```sh
pnpm install
pnpm build      # both packages
pnpm test       # vitest
pnpm typecheck
pnpm lint       # oxlint + oxfmt
```

Each script also exists inside the individual packages, where `pnpm dev` runs
`tsc --watch`.

The repository holds two published packages: `vite-plugin-spfx`, the plugin, and
`create-vite-plugin-spfx`, the scaffolder behind `npm create`.

The plugin builds with plain `tsc` on purpose: a Node-only build-time package
gains nothing from bundling, while per-file output keeps stack traces pointing
at real source lines. `inlineSources` makes the maps self-contained — run Node
with `--enable-source-maps` to see them resolved.

## Source layout

```
src/
  index.ts              build plugin — Vite hooks and orchestration
  serve.ts              dev server plugin — middlewares and watch build
  runtime.ts            browser entry — applyTheme, the only client-side code
  types.ts              public types
  _utils.ts             JSONC reading and version normalisation
  bin/
    trust-dev-cert.ts   creates and trusts the development certificate
  spfx/                 SharePoint domain knowledge — no Vite involved
    externals.ts        component discovery from node_modules
    manifest.ts         scriptResources and localized string bundles
    solution.ts         .sppkg writer
    debug.ts            serve.json, debug URL, manifests module
    certificate.ts      the Rushstack certificate store
  transform/            bundler-facing code transformations
    amd.ts              ESM → named AMD, with a no-ESM-survives assertion
    styles.ts           theme tokens (PostCSS) and style injection
```

`spfx/` is pure knowledge about SharePoint's formats and would be identical
under a different bundler; `transform/` is what the bundler hooks call.

Hook ordering is load-bearing rather than incidental: minification runs between
`renderChunk` and `generateBundle`, so a transformation that inspects the final
code has to sit on the right side of it, and the bundle object must not be
mutated where Rolldown does not expect it. Every one of those rules was learned
from a bug that passed in dev and broke in production — which is why the
contract test builds in production mode.

## Tests

Two jobs, not coverage.

**Contract** runs a real _production_ build over a fixture and asserts what the
loader requires: one named define whose id is `<componentId>_<version>`, a
dependency array equal to the `scriptResources` keys in order, no surviving ESM
syntax, no emitted stylesheet, theme fallbacks intact after minification, and a
package containing exactly the referenced assets. Production on purpose — every
bug this plugin has had passed in dev and broke in production, because
minification runs between `renderChunk` and `generateBundle`.

**Baseline** checks against `test/fixtures/spfx-<version>/`, frozen output of a
stock Heft project. Whatever Microsoft's toolchain produces is by definition
what SharePoint accepts, so those files are the contract. They are committed
rather than read from a neighbouring project: the package has to test
standalone, and a baseline that skips when a path is missing protects nothing.
Refreshing them after an SPFx upgrade is a deliberate step, and the diff is the
point — see
[the fixture README](./packages/vite-plugin-spfx/test/fixtures/spfx-1.23.2/README.md).

The suite is mutation-checked: changing a hardcoded GUID, dropping a theme
fallback, removing import deduplication or leaving `requiresCustomScript` in the
package each turn it red.

## The scaffolder's templates

`packages/create-vite-plugin-spfx/templates/` holds the five projects `npm
create` writes. They are ordinary SPFx projects with `__PLACEHOLDER__` tokens
where the identity goes — the component directory, the class, the manifest alias,
the strings specifier, the bundle key, three GUIDs and the plugin version. The
scaffolder fills those in from the web part title, in file contents and in file
and folder names alike, so a generated project is named after itself.

Editing one is editing a real project, with one catch: a template is never built
in place. Nothing type-checks it, and `vite.config.mjs` names files that only
exist after substitution. Verify by generating:

```sh
pnpm build
pnpm --filter create-vite-plugin-spfx verify-templates          # all five
pnpm --filter create-vite-plugin-spfx verify-templates solid    # just one
```

That drives the real CLI, installs the plugin from a packed tarball rather than a
workspace link — what a user gets from the registry — and fails unless a bundle
and a `.sppkg` come out. It takes a full `npm install` per template, so it runs in
CI and on demand rather than as part of `pnpm test`. On failure the generated
projects are kept and their path printed.

`pnpm test` covers the substitution itself: that no placeholder survives, that
nothing is still named after another template, and that the derived names agree
with what `vite.config.mjs` references.

Adding a framework means adding a template directory and listing it in the
`TEMPLATES` array in `src/index.ts`. Copy the closest existing template and keep
the structure identical — the point of the set is that a diff between two of them
shows only what the framework changes.

## Verifying against a tenant

Tests cover the formats; they cannot prove SharePoint accepts them. Before a
release, scaffold a project and run it both ways — `npm run dev` against the
hosted workbench and a real site page, then upload its `.sppkg` to an app
catalog. Theming, the property pane and a non-English page language are the parts
that break quietly.

Put your tenant in the scaffolded project's `.env.local`, which its `.gitignore`
already excludes:

```sh
SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com
```

## Releasing

Release both packages together and on the same version number. The templates
pin the plugin as `^<scaffolder version>`, so a scaffolder published ahead of the
plugin would generate projects whose dependency does not exist yet.

```sh
pnpm -r --filter "./packages/*" exec npm version <patch|minor|major>
pnpm --filter vite-plugin-spfx publish
pnpm --filter create-vite-plugin-spfx publish
```

`prepack` runs the build in both. `files` keeps everything else out of the
tarballs — `["dist"]` for the plugin, `["dist", "templates"]` for the scaffolder.
Check with `npm pack --dry-run`.

One npm quirk the scaffolder works around: a file named `.gitignore` is stripped
from a published tarball, so templates carry it as `_gitignore` and it is renamed
during scaffolding. A template file that must start with a dot needs the same
treatment in `src/scaffold.ts`.

## Conventions

Formatting and linting are oxfmt and oxlint; run `pnpm fmt` in the package to
apply both. Imports inside `src/` carry the `.ts` extension. The plugin must
keep working standalone — no imports that reach outside its own package, and no
test that depends on a sibling directory existing.
