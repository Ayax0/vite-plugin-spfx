import { describe, expect, it } from "vitest";
import { rewriteThemeTokens, spfxThemeTokens } from "../src/transform/styles.ts";

/** Drive the PostCSS plugin without pulling in postcss itself. */
function runPostcss(declarations: Array<{ value: string }>): void {
  const plugin = spfxThemeTokens();
  plugin.OnceExit({ walkDecls: (cb) => declarations.forEach(cb) });
}

describe("theme tokens", () => {
  it("turns a token into a custom property with its declared default", () => {
    expect(rewriteThemeTokens(`color: "[theme:bodyText, default: #323130]";`)).toBe(
      "color: var(--bodyText, #323130);",
    );
  });

  it("accepts single quotes, which Sass emits for quoted strings", () => {
    expect(rewriteThemeTokens(`color: '[theme:link, default:#03787c]';`)).toBe(
      "color: var(--link, #03787c);",
    );
  });

  it("emits a bare custom property when no default is declared", () => {
    expect(rewriteThemeTokens(`color: "[theme:link]";`)).toBe("color: var(--link);");
  });

  it("backfills the default onto a bare var() for the same slot", () => {
    // The SPFx scaffold writes the token and then overrides it with a plain
    // var(). Without a fallback that second declaration is invalid at computed
    // value time on hosts that do not expose the slot — and an invalid
    // declaration inherits instead of falling back to the previous one.
    const css = `a { color: "[theme:link, default:#03787c]"; color: var(--link); }`;
    expect(rewriteThemeTokens(css)).toBe(
      "a { color: var(--link, #03787c); color: var(--link, #03787c); }",
    );
  });

  it("leaves unrelated custom properties untouched", () => {
    expect(rewriteThemeTokens(`color: var(--brand);`)).toBe("color: var(--brand);");
  });

  it("applies the same rules through the PostCSS plugin", () => {
    // This form is the one that actually runs: it has to happen before the CSS
    // minifier drops the themed declaration as redundant.
    const declarations = [
      { value: `"[theme:bodyText, default: #323130]"` },
      { value: `var(--bodyText)` },
    ];
    runPostcss(declarations);
    expect(declarations[0]!.value).toBe("var(--bodyText, #323130)");
    expect(declarations[1]!.value).toBe("var(--bodyText, #323130)");
  });
});
