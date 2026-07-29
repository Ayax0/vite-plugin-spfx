#!/usr/bin/env node
import { certificateStore, readCertificate, trustDevCertificate } from "../spfx/certificate.ts";

/**
 * `npx spfx-trust-dev-cert`
 *
 * Creates the local development certificate and adds it to the OS trust store,
 * without needing Heft. Uses the same store SPFx's tooling does, so it is
 * interchangeable with `heft trust-dev-cert`.
 */
const store = certificateStore(process.cwd());

if (readCertificate(store)) {
  console.log(`Certificate already present in ${store} — nothing to do.`);
  process.exit(0);
}

console.log(`Creating a development certificate in ${store}.`);
console.log("Your OS will ask for permission to trust it.\n");

try {
  await trustDevCertificate();
  console.log("\nDone. `vite dev` will pick it up.");
} catch (error) {
  console.error(`\n${error instanceof Error ? error.message : String(error)}`);
  process.exit(1);
}
