# SPFx 1.23.2 baseline

Frozen output of Microsoft's own toolchain — a stock `yo @microsoft/sharepoint`
React web part built with Heft and webpack. Whatever it produces is, by
definition, what SharePoint accepts, so these files are the contract our output
has to stay compatible with.

Committed rather than read from a sibling project on purpose: this package must
test standalone, and a baseline that silently disappears when the reference is
missing protects nothing.

These files are excluded from linting and formatting — they are Microsoft's
output, and rewriting them would defeat the point of freezing them.

| File | Origin |
| --- | --- |
| `component.manifest.json` | `dist/<componentId>.manifest.json` |
| `solution-parts.json` | file list of `sharepoint/solution/debug/`, GUIDs replaced with `<feature>` / `<component>`, `ClientSideAssets/` contents omitted |

## Refreshing after an SPFx upgrade

```sh
npx yo @microsoft/sharepoint --solution-name spfx-reference \
    --component-type webpart --framework react --skip-feature-deployment
cd spfx-reference && npm install && npm run build
```

Then copy `dist/<guid>.manifest.json` over `component.manifest.json`, and rebuild
`solution-parts.json` from `sharepoint/solution/debug/`.

The diff is the point: it shows exactly what Microsoft changed. Failing tests
after a refresh mean the plugin needs to follow — a moved component GUID, a new
manifest field, an added OPC part.
