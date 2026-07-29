import { defineConfig } from "vite";
import { svelte, vitePreprocess } from "@sveltejs/vite-plugin-svelte";
import { spfx } from "vite-plugin-spfx";

export default defineConfig({
  // Like Vue, Svelte is not a SharePoint component and gets bundled — though
  // Svelte compiles away, so the runtime it contributes is small.
  plugins: [
    // vitePreprocess is what makes <style lang="scss"> work; without it the
    // SCSS reaches PostCSS unprocessed.
    svelte({ preprocess: vitePreprocess() }),
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
