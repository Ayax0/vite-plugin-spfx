import { describe, expect, it } from "vitest";
import { init, parse } from "es-module-lexer";
import { esmToAmd } from "../src/transform/amd.ts";

const ID = "abc_1.0.0";

/**
 * The SharePoint loader receives a classic script, so anything the conversion
 * leaves behind is a syntax error at runtime — in production only, because dev
 * builds are unminified and happen to survive some of these mistakes.
 */
async function assertNoEsm(code: string): Promise<void> {
  await init;
  const [imports, exports] = parse(code);
  expect(imports.filter((i) => i.n !== undefined)).toHaveLength(0);
  expect(exports).toHaveLength(0);
}

describe("esmToAmd", () => {
  it("emits a named define whose id is the component id and version", async () => {
    const { code } = await esmToAmd(`const a = 1;\nexport { a as default };\n`, { amdId: ID });
    expect(code).toContain(`define("abc_1.0.0", []`);
    expect(code).toContain(`return { __esModule: true, "default": a };`);
  });

  it("handles minified input", async () => {
    // The regression that shipped a broken bundle: a line-anchored regex found
    // `export{...}` only when it sat on its own line. Minifiers put it at the
    // end of a very long one.
    const minified = `import{b as c}from"x";import*as d from"y";const e=c(d);export{e as default};`;
    const { code, deps } = await esmToAmd(minified, { amdId: ID });
    expect(deps).toEqual(["x", "y"]);
    await assertNoEsm(code);
    expect(code).toContain(`return { __esModule: true, "default": e };`);
  });

  it("binds every import form onto the factory parameters", async () => {
    const source = [
      `import * as ns from "a";`,
      `import def from "b";`,
      `import { one, two as alias } from "c";`,
      `export { ns as default };`,
    ].join("\n");
    const { code, deps } = await esmToAmd(source, { amdId: ID });

    expect(deps).toEqual(["a", "b", "c"]);
    expect(code).toContain("var ns = __spfx_dep_0;");
    // Default imports need interop: an AMD module is not an ES namespace.
    expect(code).toContain(
      `var def = (__spfx_dep_1 && __spfx_dep_1.__esModule) ? __spfx_dep_1["default"] : __spfx_dep_1;`,
    );
    expect(code).toContain(`var one = __spfx_dep_2["one"];`);
    expect(code).toContain(`var alias = __spfx_dep_2["two"];`);
  });

  it("collapses repeated specifiers into one dependency", async () => {
    // Order and arity of the define array must match scriptResources exactly;
    // a duplicated entry would shift every following factory argument.
    const source = `import { a } from "x";\nimport { b } from "x";\nexport { a as default };`;
    const { code, deps } = await esmToAmd(source, { amdId: ID });
    expect(deps).toEqual(["x"]);
    expect(code).toContain(`var a = __spfx_dep_0["a"];`);
    expect(code).toContain(`var b = __spfx_dep_0["b"];`);
  });

  it("preserves import order, because the loader maps deps positionally", async () => {
    const source = `import "z";\nimport "a";\nimport "m";\nconst v = 1;\nexport { v as default };`;
    const { deps } = await esmToAmd(source, { amdId: ID });
    expect(deps).toEqual(["z", "a", "m"]);
  });

  it("keeps a prelude inside the factory", async () => {
    const { code } = await esmToAmd(`const a = 1;\nexport { a as default };`, {
      amdId: ID,
      prelude: ["/*marker*/"],
    });
    const factoryStart = code.indexOf('"use strict";');
    expect(code.indexOf("/*marker*/")).toBeGreaterThan(factoryStart);
  });

  it("captures the script URL for asset resolution", async () => {
    const { code } = await esmToAmd(`const a = 1;\nexport { a as default };`, { amdId: ID });
    // Equivalent of SPFx's SetPublicPathCurrentScriptPlugin — must be read at
    // script evaluation time, not inside the factory, which runs later.
    expect(code).toContain("document.currentScript");
    expect(code.indexOf("__spfx_base")).toBeLessThan(code.indexOf("define("));
  });

  it("rejects dynamic import instead of emitting a chunk the loader cannot fetch", async () => {
    await expect(esmToAmd(`const p = import("x");`, { amdId: ID })).rejects.toThrow(
      /dynamic import/,
    );
  });

  it("rejects an export it cannot resolve to a local binding", async () => {
    await expect(esmToAmd(`export default 42;`, { amdId: ID })).rejects.toThrow(
      /local binding|export statement/,
    );
  });
});
