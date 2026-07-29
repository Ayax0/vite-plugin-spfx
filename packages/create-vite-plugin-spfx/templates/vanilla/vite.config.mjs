import { defineConfig } from "vite";
import { spfx } from "vite-plugin-spfx";

// No framework plugin, and no framework: the whole config is the spfx() call.
// SharePoint only asks for a class that renders into a DOM element, so a web
// part without any framework at all is a legitimate — and by far the smallest —
// way to satisfy that.
export default defineConfig({
  plugins: [
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
