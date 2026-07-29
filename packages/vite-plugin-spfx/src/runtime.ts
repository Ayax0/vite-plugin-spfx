/**
 * Browser-side helpers, imported by the component — never by the build.
 *
 *     import { applyTheme } from 'vite-plugin-spfx/runtime'
 */

/** The shape SPFx hands to `onThemeChanged`. */
export interface SpfxTheme {
  palette?: Readonly<Record<string, string | undefined>>;
  semanticColors?: Readonly<Record<string, string | undefined>>;
}

/**
 * Publish the host theme as CSS custom properties on the component root.
 *
 * The plugin rewrites `"[theme:bodyText, default:#323130]"` to
 * `var(--bodyText, #323130)` at build time, but something has to actually set
 * `--bodyText`. Stock SPFx makes every web part hand-write one
 * `style.setProperty()` per slot inside `onThemeChanged` — easy to forget (the
 * styles then silently fall back to their defaults) and easy to leave incomplete
 * once the CSS uses a new slot.
 *
 * This publishes every slot the host provides, so the CSS decides which matter.
 */
export function applyTheme(element: HTMLElement | undefined, theme: SpfxTheme | undefined): void {
  if (!element || !theme) return;
  // semanticColors last: the more specific layer wins on overlap.
  for (const slots of [theme.palette, theme.semanticColors]) {
    if (!slots) continue;
    for (const [name, value] of Object.entries(slots)) {
      if (typeof value === "string") element.style.setProperty(`--${name}`, value);
    }
  }
}
