import MagicString from "magic-string";
import { init as initLexer, parse as parseEsm } from "es-module-lexer";

export interface AmdResult {
  code: string;
  map: ReturnType<MagicString["generateMap"]>;
  /** Dependency names in `define()` order — must match the scriptResources keys. */
  deps: string[];
}

interface ImportClause {
  namespace?: string;
  defaultName?: string;
  named: Array<[imported: string, local: string]>;
}

/**
 * Parse the binding clause of a generated import statement.
 * Whitespace is optional everywhere: minified output emits `import{a}from"x"`.
 */
function parseImportClause(statement: string): ImportClause {
  const result: ImportClause = { named: [] };
  const match = /^import\s*([\s\S]*?)\s*from\s*['"]/.exec(statement);
  if (!match) return result; // bare `import 'x'`, side effects only

  let clause = match[1]!.trim();
  const namespaceOnly = /^\*\s*as\s+([A-Za-z_$][\w$]*)$/.exec(clause);
  if (namespaceOnly) return { ...result, namespace: namespaceOnly[1]! };

  if (!clause.startsWith("{")) {
    const comma = clause.indexOf(",");
    if (comma === -1) return { ...result, defaultName: clause };
    result.defaultName = clause.slice(0, comma).trim();
    clause = clause.slice(comma + 1).trim();
    const namespace = /^\*\s*as\s+([A-Za-z_$][\w$]*)$/.exec(clause);
    if (namespace) return { ...result, namespace: namespace[1]! };
  }

  if (clause.startsWith("{")) {
    for (const raw of clause.slice(1, clause.lastIndexOf("}")).split(",")) {
      const part = raw.trim();
      if (!part) continue;
      const aliased = /^(.+?)\s+as\s+(.+)$/.exec(part);
      result.named.push(aliased ? [aliased[1]!.trim(), aliased[2]!.trim()] : [part, part]);
    }
  }
  return result;
}

/**
 * Rewrite an ESM chunk into the named AMD module the SharePoint loader expects:
 *
 *     define("<componentId>_<version>", [<scriptResources keys>], factory)
 *
 * Neither Rolldown nor Vite 8 can emit AMD, so the wrapper is synthesised here.
 * The header also captures `document.currentScript` at evaluation time, which is
 * how asset URLs resolve — the equivalent of SPFx's SetPublicPathCurrentScriptPlugin.
 *
 * Call this from `renderChunk`: the content hash is derived from its result, so
 * transforming later would leave the hash blind to the wrapper.
 */
export async function esmToAmd(
  code: string,
  { amdId, prelude = [] }: { amdId: string; prelude?: string[] },
): Promise<AmdResult> {
  await initLexer;
  const [imports, exports] = parseEsm(code);
  const source = new MagicString(code);

  const deps: string[] = [];
  const params: string[] = [];
  const paramOf = new Map<string, string>();
  const bindings: string[] = [];

  for (const statement of imports) {
    if (statement.d > -1) {
      throw new Error(`[spfx] dynamic import() is not supported yet (found "${statement.n}")`);
    }
    if (statement.n === undefined) continue; // import.meta

    let end = statement.se;
    while (end < code.length && (code[end] === ";" || code[end] === " ")) end++;
    source.remove(statement.ss, end);

    let param = paramOf.get(statement.n);
    if (!param) {
      param = `__spfx_dep_${deps.length}`;
      deps.push(statement.n);
      params.push(param);
      paramOf.set(statement.n, param);
    }

    const clause = parseImportClause(code.slice(statement.ss, statement.se));
    if (clause.namespace) bindings.push(`var ${clause.namespace} = ${param};`);
    if (clause.defaultName) {
      bindings.push(
        `var ${clause.defaultName} = (${param} && ${param}.__esModule) ? ${param}["default"] : ${param};`,
      );
    }
    for (const [imported, local] of clause.named) {
      bindings.push(`var ${local} = ${param}[${JSON.stringify(imported)}];`);
    }
  }

  // Rebuild the exports as a return value. Locating the statements with a
  // line-anchored regex is wrong: minified output puts `export{x as default}` at
  // the end of a long line, where it survives and the loader dies with
  // "Unexpected token 'export'". Derive the bounds from the lexer instead.
  const returned: string[] = [];
  const statements = new Map<number, number>();

  for (const entry of exports) {
    if (!entry.ln) {
      throw new Error(
        `[spfx] export "${entry.n}" has no resolvable local binding — cannot convert to AMD`,
      );
    }
    returned.push(`${JSON.stringify(entry.n)}: ${entry.ln}`);

    const head = code.lastIndexOf("export", Math.min(entry.s, entry.ls ?? entry.s));
    const close = code.indexOf("}", Math.max(entry.e, entry.le ?? entry.e));
    if (head === -1 || close === -1) {
      throw new Error(`[spfx] could not locate the export statement for "${entry.n}"`);
    }
    let end = close + 1;
    while (end < code.length && (code[end] === ";" || code[end] === " " || code[end] === "\t"))
      end++;
    statements.set(head, end);
  }
  for (const [start, end] of statements) source.remove(start, end);

  // The body must be free of ESM syntax before it goes inside define(). Without
  // this check a missed statement only surfaces at runtime in SharePoint, and
  // only in production builds.
  const [leftoverImports, leftoverExports] = parseEsm(source.toString());
  if (leftoverImports.length > 0 || leftoverExports.length > 0) {
    throw new Error(
      `[spfx] ESM->AMD conversion is incomplete: ${leftoverImports.length} import(s) and ` +
        `${leftoverExports.length} export(s) survived. The SharePoint loader would fail ` +
        `with "Unexpected token 'export'".`,
    );
  }

  source.prepend(
    `(function () {\n` +
      `var __spfx_script = document.currentScript;\n` +
      `var __spfx_base = String((__spfx_script && __spfx_script.src) || '').replace(/[^/]*$/, '');\n` +
      `define(${JSON.stringify(amdId)}, ${JSON.stringify(deps)}, function (${params.join(", ")}) {\n` +
      `"use strict";\n` +
      [...bindings, ...prelude].join("\n") +
      "\n",
  );
  source.append(`\nreturn { __esModule: true, ${returned.join(", ")} };\n});\n})();\n`);

  return { code: source.toString(), map: source.generateMap({ hires: true }), deps };
}
