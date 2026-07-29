import { defineConfig } from "vite";
import solid from "vite-plugin-solid";
import { spfx } from "vite-plugin-spfx";

export default defineConfig({
  // Solid is not a SharePoint component and gets bundled. Like Svelte it
  // compiles JSX away into direct DOM calls, so what it contributes at runtime
  // is a small reactivity core rather than a virtual DOM.
  plugins: [
    solid(),
    spfx({
      bundleName: "__BUNDLE_NAME__",
      entry: "src/webparts/__COMPONENT_DIR__/__ALIAS__.ts",
      manifest: "src/webparts/__COMPONENT_DIR__/__ALIAS__.manifest.json",
      // Each locale becomes a standalone AMD module; the manifest gets a
      // localizedPath entry so SharePoint picks one by page language.
      localizedResources: {
        __STRINGS__: "src/webparts/__COMPONENT_DIR__/loc/{locale}.js",
      },
    }),
  ],
  build: {
    // Only here to keep the logo a real file, so the example demonstrates the
    // runtime asset URL. Not an SPFx requirement — Vite's default (inline below
    // 4 kB) works fine.
    assetsInlineLimit: 0,
  },
});
