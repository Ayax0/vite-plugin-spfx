import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { certificateStore, readCertificate } from "../src/spfx/certificate.ts";

function project(override?: unknown, depth = 0): string {
  const base = fs.mkdtempSync(path.join(os.tmpdir(), "spfx-cert-"));
  if (override !== undefined) {
    fs.mkdirSync(path.join(base, ".vscode"), { recursive: true });
    fs.writeFileSync(
      path.join(base, ".vscode/debug-certificate-manager.json"),
      JSON.stringify(override),
    );
  }
  // Nested project directories still find an override further up the tree.
  const root = path.join(base, ...Array.from({ length: depth }, (_, i) => `level${i}`));
  fs.mkdirSync(root, { recursive: true });
  return root;
}

describe("certificateStore", () => {
  it("defaults to the store SPFx tooling uses", () => {
    expect(certificateStore(project())).toBe(path.join(os.homedir(), ".rushstack"));
  });

  it("honours a storePath override, like the certificate manager does", () => {
    const root = project({ storePath: "./certs" });
    expect(certificateStore(root)).toBe(path.join(root, "certs"));
  });

  it("expands a leading ~ to the home directory", () => {
    const root = project({ storePath: "~/my-certs" });
    expect(certificateStore(root)).toBe(path.join(os.homedir(), "my-certs"));
  });

  it("finds an override in a parent directory", () => {
    const root = project({ storePath: "~/parent-certs" }, 3);
    expect(certificateStore(root)).toBe(path.join(os.homedir(), "parent-certs"));
  });

  it("ignores an override that sets no storePath", () => {
    expect(certificateStore(project({}))).toBe(path.join(os.homedir(), ".rushstack"));
  });
});

describe("readCertificate", () => {
  it("returns nothing when the store is empty", () => {
    expect(readCertificate(project())).toBeUndefined();
  });

  it("needs both halves, not just one", () => {
    const store = project();
    fs.writeFileSync(path.join(store, "rushstack-serve.pem"), "cert");
    expect(readCertificate(store)).toBeUndefined();
  });

  it("reads the pair the way the dev server expects", () => {
    const store = project();
    fs.writeFileSync(path.join(store, "rushstack-serve.key"), "KEY");
    fs.writeFileSync(path.join(store, "rushstack-serve.pem"), "CERT");
    const certificate = readCertificate(store);
    expect(certificate?.key.toString()).toBe("KEY");
    expect(certificate?.cert.toString()).toBe("CERT");
  });
});
