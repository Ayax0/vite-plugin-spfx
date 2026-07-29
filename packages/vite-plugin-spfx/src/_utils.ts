import fs from "node:fs";
import { parse, printParseErrorCode, type ParseError } from "jsonc-parser";

/** SPFx config and manifest files are JSONC — they carry `//` comments. */
export function readJsonc<T>(file: string): T {
  const errors: ParseError[] = [];
  const value = parse(fs.readFileSync(file, "utf8"), errors, { allowTrailingComma: true });
  if (errors.length > 0) {
    const [first] = errors;
    throw new Error(
      `[spfx] ${file} is not valid JSONC: ${printParseErrorCode(first!.error)} at offset ${first!.offset}`,
    );
  }
  return value as T;
}

/** SPFx strips any prerelease suffix from versions used in manifests. */
export const cleanVersion = (version: string): string => String(version).split("-")[0]!;
