/**
 * SPFx authors themed styles as `color: "[theme:bodyText, default: #323130]"`.
 * That is an invalid CSS value; stock SPFx substitutes it at runtime with
 * @microsoft/load-themed-styles. We rewrite it to a CSS custom property instead,
 * which modern hosts resolve natively — no runtime, no extra bytes.
 *
 * `applyTheme` (see runtime.ts) is what publishes those custom properties.
 */

const THEME_TOKEN = /["']\[theme:\s*([\w-]+)\s*(?:,\s*default:\s*([^\]]*?))?\s*\]["']/g;
const BARE_VAR = /var\(\s*--([\w-]+)\s*\)/g;

/** Minimal structural types, so postcss does not become a dependency. */
interface Declaration {
  value: string;
}
interface Root {
  walkDecls(callback: (decl: Declaration) => void): void;
}

function rewrite(value: string, defaults: Map<string, string>): string {
  return value.replace(THEME_TOKEN, (_all, slot: string, fallback?: string) => {
    if (fallback) defaults.set(slot, fallback.trim());
    return fallback ? `var(--${slot}, ${fallback.trim()})` : `var(--${slot})`;
  });
}

/**
 * The SPFx scaffold pairs every theme token with a bare `var(--slot)` on the
 * next line, which wins by cascade order. Without a fallback that declaration is
 * invalid-at-computed-value-time on hosts that do not expose the slot — and an
 * invalid declaration inherits rather than falling back to the previous one.
 */
function backfillDefaults(value: string, defaults: Map<string, string>): string {
  return value.replace(BARE_VAR, (all, slot: string) => {
    const fallback = defaults.get(slot);
    return fallback ? `var(--${slot}, ${fallback})` : all;
  });
}

/**
 * PostCSS plugin form. This MUST run before CSS minification: the minifier drops
 * the themed declaration as redundant (the `var()` line overrides it), so a
 * rewrite running afterwards never sees the token and loses every fallback.
 */
export function spfxThemeTokens() {
  return {
    postcssPlugin: "spfx-theme-tokens",
    OnceExit(root: Root) {
      const defaults = new Map<string, string>();
      root.walkDecls((decl) => {
        decl.value = rewrite(decl.value, defaults);
      });
      root.walkDecls((decl) => {
        decl.value = backfillDefaults(decl.value, defaults);
      });
    },
  };
}
spfxThemeTokens.postcss = true;

/** String form — a safety net for CSS that never passed through PostCSS. */
export function rewriteThemeTokens(css: string): string {
  const defaults = new Map<string, string>();
  return backfillDefaults(rewrite(css, defaults), defaults)
    .replace(/\/\*\$vite\$:\d+\*\//g, "") // internal Vite marker, not ours to ship
    .trim();
}

/**
 * Injected into the AMD wrapper with a placeholder token instead of the real CSS,
 * because Vite's finished stylesheet only exists in generateBundle while
 * minification happens before it. The token is a STRING literal, not a comment:
 * the minifier strips comments, which would silently drop every style.
 *
 * Substituting an opaque token after rendering is the technique Vite itself uses
 * for assets (`__VITE_ASSET__<hash>__`).
 */
export const STYLE_TOKEN = "__SPFX_STYLES_TOKEN__";
export const STYLE_TOKEN_PATTERN = new RegExp(`(["'\`])${STYLE_TOKEN}\\1`);

export function styleInjector(marker: string): string {
  return (
    `(function(__c){if(!__c){return;}` +
    `var __s=document.createElement("style");` +
    `__s.setAttribute("data-spfx-vite",${JSON.stringify(marker)});` +
    `var __n=(typeof window!=="undefined"&&window.CSPSettings&&window.CSPSettings.nonce);` +
    `if(__n){__s.setAttribute("nonce",__n);}` +
    `__s.appendChild(document.createTextNode(__c));` +
    `document.head.appendChild(__s);})(${JSON.stringify(STYLE_TOKEN)});`
  );
}
