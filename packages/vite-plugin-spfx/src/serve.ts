import fs from "node:fs";
import path from "node:path";
import { styleText } from "node:util";
import { loadEnv, type Plugin, type ViteDevServer } from "vite";
import type { ServerResponse } from "node:http";

import { certificateStore, readCertificate } from "./spfx/certificate.ts";
import {
  debugQuery,
  LEGACY_MANIFESTS_PATH,
  MANIFESTS_PATH,
  manifestsModule,
  readServeConfig,
  resolveInitialPage,
} from "./spfx/debug.ts";
import { displayNameFor } from "./spfx/solution.ts";
import type { ComponentManifest, SpfxOptions } from "./types.ts";

/**
 * Dev server for SharePoint.
 *
 * Vite's dev server serves native ESM; the SharePoint loader is a RequireJS fork
 * that cannot consume it. So this uses the dev server for what it is good at —
 * an HTTPS host with middleware, config resolution and a process to live in —
 * and runs a watch build alongside to produce the AMD bundle SharePoint needs.
 *
 * The same shape browser-extension and Electron plugins use, and for the same
 * reason: the target cannot eat a module graph.
 */

const MIME: Record<string, string> = {
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".map": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
};

const RELOAD_PATH = "/__spfx_reload";

/**
 * The bundle is fetched from the SharePoint origin, so every response is
 * cross-origin, and Chromium's Local Network Access needs the extra header.
 *
 * Set per response rather than through `server.headers`: plugin middlewares run
 * before Vite's own, so a response written here never reaches the header
 * middleware.
 */
const CORS_HEADERS: Record<string, string> = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "HEAD, GET, OPTIONS",
  "Access-Control-Allow-Private-Network": "true",
};

function respond(response: ServerResponse, body: string | Buffer, contentType: string): void {
  response.writeHead(200, {
    ...CORS_HEADERS,
    "Content-Type": contentType,
    "Cache-Control": "no-store",
  });
  response.end(body);
}

function missingCertificate(store: string): string {
  return (
    `[spfx] no development certificate in ${store}.\n\n` +
    `SharePoint pages are https and load the bundle cross-origin, so an http\n` +
    `dev server is blocked as mixed content and an untrusted certificate is\n` +
    `rejected outright — browsers offer no click-through for subresources.\n\n` +
    `Create a trusted one:      npx spfx-trust-dev-cert\n` +
    `Already use Heft?          it is the same store, nothing more to do\n` +
    `Bring your own:            set server.https in vite.config`
  );
}

/** Env var SPFx itself reads, so existing setups keep working. */
const TENANT_VAR = "SPFX_SERVE_TENANT_DOMAIN";

const dim = (text: string): string => styleText("dim", text);

export function spfxServe(options: SpfxOptions): Plugin {
  let root: string;
  let outDir: string;
  let port: number;
  let debugPage: string | undefined;
  const clients = new Set<ServerResponse>();

  const readManifests = (): ComponentManifest[] =>
    fs.existsSync(outDir)
      ? fs
          .readdirSync(outDir)
          .filter((file) => file.endsWith(".manifest.json"))
          .map((file) => JSON.parse(fs.readFileSync(path.join(outDir, file), "utf8")))
      : [];

  /** File names of the entry bundles, read back from the emitted manifests. */
  const entryFiles = (): string[] =>
    readManifests()
      .map((manifest) => {
        const config = manifest.loaderConfig;
        const entry = config?.scriptResources[config.entryModuleId];
        return entry && entry.type === "path" ? entry.path : undefined;
      })
      .filter((name): name is string => Boolean(name));

  return {
    name: "vite-plugin-spfx:serve",
    apply: "serve",

    config(userConfig, env) {
      root = path.resolve(userConfig.root ?? process.cwd());
      const serve = readServeConfig(root);
      port = userConfig.server?.port ?? serve.port;

      // Fail here rather than serving http: the failure mode downstream is a
      // silently blank web part with nothing in the console pointing at TLS.
      let certificate;
      if (serve.https && !userConfig.server?.https) {
        const store = certificateStore(root);
        certificate = readCertificate(store);
        if (!certificate) throw new Error(missingCertificate(store));
      }

      // The tenant is per-developer data, so it comes from the environment
      // rather than committed config. loadEnv picks up `.env.local` too, which
      // is the idiomatic Vite place for machine-specific values and is
      // gitignored by convention — no shell export needed.
      const dotenv = loadEnv(env.mode, root, "SPFX_");
      const tenantDomain = process.env[TENANT_VAR] ?? dotenv[TENANT_VAR];
      const page = resolveInitialPage(options.serveInitialPage ?? serve.initialPage, tenantDomain);
      const scheme = certificate || userConfig.server?.https ? "https" : "http";
      debugPage = page
        ? page + debugQuery(`${scheme}://localhost:${port}${MANIFESTS_PATH}`)
        : undefined;

      return {
        server: {
          port,
          // The client gate rejects any host other than localhost/127.0.0.1.
          host: "127.0.0.1",
          strictPort: true,
          https: userConfig.server?.https ?? certificate,
          cors: true,
          // Also on Vite's own responses; our middlewares set them themselves.
          headers: { ...CORS_HEADERS, ...userConfig.server?.headers },
          // `vite dev --open` should land on the workbench, not on a Vite page
          // that means nothing here.
          open: userConfig.server?.open === true ? (debugPage ?? false) : userConfig.server?.open,
        },
      };
    },

    configResolved(resolved) {
      outDir = path.resolve(resolved.root, resolved.build.outDir);
    },

    configureServer(server) {
      server.middlewares.use(RELOAD_PATH, (request, response) => {
        // No `Connection: keep-alive` — the dev server speaks HTTP/2, where
        // connection-specific headers are illegal and Node warns about them.
        response.writeHead(200, {
          ...CORS_HEADERS,
          "Content-Type": "text/event-stream",
          "Cache-Control": "no-store",
        });
        response.write("retry: 2000\n\n");
        clients.add(response);
        request.on("close", () => clients.delete(response));
      });

      for (const route of [MANIFESTS_PATH, LEGACY_MANIFESTS_PATH]) {
        server.middlewares.use(route, (_request, response) => {
          respond(response, manifestsModule(readManifests()), MIME[".js"]!);
        });
      }

      server.middlewares.use("/dist", (request, response, next) => {
        const name = decodeURI((request.url ?? "/").split("?")[0]!).replace(/^\//, "");
        const file = path.join(outDir, name);
        if (!file.startsWith(outDir) || !fs.existsSync(file)) return next();

        let body = fs.readFileSync(file);
        // Only the entry bundle gets the reload client — one EventSource per page.
        if (entryFiles().includes(path.basename(file))) {
          body = Buffer.concat([body, Buffer.from(reloadClient(port))]);
        }
        respond(response, body, MIME[path.extname(file)] ?? "application/octet-stream");
      });

      // The watch build is what actually produces the bundle. It runs the same
      // plugin pipeline, so dev and release output cannot drift apart.
      server.httpServer?.once("listening", () => void startWatch(server, clients));
    },
  };

  function reloadClient(servePort: number): string {
    return (
      `\n/* vite-plugin-spfx live reload */\n` +
      `(function () { try {` +
      `var es = new EventSource("https://localhost:${servePort}${RELOAD_PATH}");` +
      `es.addEventListener("reload", function () { location.reload(); });` +
      `} catch (e) {} })();\n`
    );
  }

  async function startWatch(server: ViteDevServer, sseClients: Set<ServerResponse>): Promise<void> {
    const { build } = await import("vite");
    const { configFile } = server.config;

    if (!configFile) {
      server.config.logger.error(
        "[spfx] the dev server needs a Vite config file to run its watch build.",
      );
      return;
    }

    const watcher = await build({
      root,
      configFile,
      mode: "development",
      logLevel: "warn",
      build: { watch: {} },
    });

    if (!("on" in watcher)) return;

    let first = true;
    watcher.on("event", (event) => {
      if (event.code === "BUNDLE_END") {
        // The banner needs the manifests, which only exist after a build.
        if (first) {
          first = false;
          printBanner(server);
        } else {
          server.config.logger.info(
            `${styleText("green", "spfx rebuilt")} ${dim(`in ${event.duration}ms`)}`,
            { timestamp: true },
          );
        }
        for (const client of sseClients) client.write("event: reload\ndata: 1\n\n");
      }
      if (event.code === "ERROR") server.config.logger.error(String(event.error));
    });
  }

  /**
   * Printed in Vite's own shape, so the SPFx lines read as part of the banner
   * rather than as something bolted on. Node's styleText drops the colours by
   * itself when the output is piped or NO_COLOR is set.
   */
  function printBanner(server: ViteDevServer): void {
    const scheme = server.config.server.https ? "https" : "http";
    const query = debugQuery(`${scheme}://localhost:${port}${MANIFESTS_PATH}`);
    const components = readManifests().map(displayNameFor);

    const line = (label: string, value: string): string =>
      `${styleText(["green", "bold"], "➜")}  ${styleText("bold", `${label}:`.padEnd(11))}${value}`;

    const lines = [""];
    if (components.length > 0) {
      lines.push(line("Component", components.map((name) => styleText("cyan", name)).join(", ")));
    }
    // The link is there to be clicked; the query to be pasted onto some other
    // page. Both, because they are used differently.
    if (debugPage) lines.push(line("Open", styleText("cyan", debugPage)));
    lines.push(line("Query", dim(query)));
    if (!debugPage) {
      lines.push(`${dim(`append it to any modern page — or set ${TENANT_VAR} in .env.local`)}`);
    }
    lines.push("");

    server.config.logger.info(lines.join("\n"));
  }
}
