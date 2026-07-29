# __TITLE__

A SharePoint Framework web part built with [Vite](https://vite.dev) and
[vite-plugin-spfx](https://github.com/Ayax0/vite-plugin-spfx), using vue.

## Develop

```sh
npm run dev
```

This starts an HTTPS server and a watch build together, then prints a link to
open in SharePoint.

Two things it needs from you:

**Your tenant.** Put it in `.env.local` (gitignored, so it stays yours):

```sh
SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com
```

**A trusted certificate.** SharePoint pages are https and load your bundle
cross-origin, so an untrusted certificate is rejected outright. If this machine
has ever run the SPFx tooling you already have one. Otherwise:

```sh
npx spfx-trust-dev-cert
```

## Ship

```sh
npm run build
```

Writes `dist/`, `release/` and `sharepoint/solution/__PROJECT_NAME__.sppkg`.
Upload the `.sppkg` to your tenant app catalog, approve it, then add the web
part from the toolbox.

## Where things are

```
config/package-solution.json      what the .sppkg contains
config/serve.json                port and workbench URL
src/webparts/__COMPONENT_DIR__/  the component, its manifest, strings and assets
vite.config.mjs                  the spfx() plugin options
```

Everything is named after the web part, so nothing has to be renamed before you
start: the class is `__ALIAS__`, the strings resolve under `__STRINGS__`, and
the build writes `__BUNDLE_NAME__.js`.

Full documentation:
https://github.com/Ayax0/vite-plugin-spfx/tree/main/packages/vite-plugin-spfx#readme
