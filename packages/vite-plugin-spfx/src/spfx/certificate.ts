import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { readJsonc } from "../_utils.ts";

/**
 * The development certificate SharePoint work uses.
 *
 * It has to be *trusted*, not merely present: the bundle is loaded from the
 * SharePoint origin as a cross-origin subresource, and browsers block those on
 * an untrusted certificate without offering a click-through. That rules out
 * self-signed throwaways such as @vitejs/plugin-basic-ssl.
 *
 * We use the same store SPFx's own tooling uses, so a machine set up for Heft
 * needs no second certificate — and `spfx-trust-dev-cert` creates one for
 * machines that were never set up for Heft at all.
 */

export interface Certificate {
  key: Buffer;
  cert: Buffer;
}

/**
 * `~/.rushstack`, unless a `.vscode/debug-certificate-manager.json` up the tree
 * points elsewhere — the override the certificate manager itself honours.
 */
export function certificateStore(root: string): string {
  for (let dir = root; ; dir = path.dirname(dir)) {
    const override = path.join(dir, ".vscode", "debug-certificate-manager.json");
    if (fs.existsSync(override)) {
      const { storePath } = readJsonc<{ storePath?: string }>(override);
      if (storePath) {
        return storePath.startsWith("~")
          ? path.join(os.homedir(), storePath.slice(1))
          : path.resolve(dir, storePath);
      }
    }
    if (dir === path.dirname(dir)) break;
  }
  return path.join(os.homedir(), ".rushstack");
}

export function readCertificate(store: string): Certificate | undefined {
  const key = path.join(store, "rushstack-serve.key");
  const cert = path.join(store, "rushstack-serve.pem");
  return fs.existsSync(key) && fs.existsSync(cert)
    ? { key: fs.readFileSync(key), cert: fs.readFileSync(cert) }
    : undefined;
}

/**
 * Create the certificate and add it to the OS trust store.
 *
 * Loaded lazily: installing a local CA is a deliberate, privileged act, so the
 * dependency stays out of the default install and out of CI.
 */
interface CertificateManagerModule {
  CertificateManager: new () => {
    ensureCertificateAsync(canGenerateNewCertificate: boolean, terminal: unknown): Promise<unknown>;
  };
}

interface TerminalModule {
  Terminal: new (provider: unknown) => unknown;
  ConsoleTerminalProvider: new () => unknown;
}

export async function trustDevCertificate(): Promise<void> {
  // Indirect specifiers on purpose: these packages are optional, so neither the
  // type checker nor a bundler should demand them. The interfaces above are the
  // whole API surface we touch.
  const managerModule = "@rushstack/debug-certificate-manager";
  const terminalModule = "@rushstack/terminal";

  let CertificateManager: CertificateManagerModule["CertificateManager"];
  let Terminal: TerminalModule["Terminal"];
  let ConsoleTerminalProvider: TerminalModule["ConsoleTerminalProvider"];

  try {
    ({ CertificateManager } = (await import(managerModule)) as CertificateManagerModule);
    ({ Terminal, ConsoleTerminalProvider } = (await import(terminalModule)) as TerminalModule);
  } catch {
    throw new Error(
      "Creating a trusted certificate needs two more packages:\n\n" +
        "  pnpm add -D @rushstack/debug-certificate-manager @rushstack/terminal\n\n" +
        "They are not a dependency of vite-plugin-spfx because installing a local\n" +
        "certificate authority is a one-off, privileged step — not something every\n" +
        "install or CI run should carry.",
    );
  }

  const terminal = new Terminal(new ConsoleTerminalProvider());
  await new CertificateManager().ensureCertificateAsync(true, terminal);
}
