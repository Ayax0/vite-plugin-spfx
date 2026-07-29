import { createHash } from "node:crypto";
import { zipSync } from "fflate";
import type { ComponentManifest, PackagedAsset, Solution, SolutionFeature } from "../types.ts";

/**
 * `.sppkg` writer — a clean-room reimplementation of Heft's package-solution phase.
 *
 * An .sppkg is an OPC (Open Packaging Conventions) zip. Byte parity with Heft is
 * not required; what SharePoint actually contracts on is:
 *   - the namespaces and relationship type URIs
 *   - element and attribute names
 *   - a leading "/" on every relationship target
 *   - the exact pseudo URL below when assets ship inside the package
 *   - [Content_Types].xml covering every extension present in the zip
 *   - scriptResources paths matching the real ClientSideAssets/ file names
 *
 * The component manifest exists in the package ONLY as the escaped
 * ComponentManifest attribute of <ClientSideComponent> — never as a loose file.
 */

const NS = {
  relationships: "http://schemas.openxmlformats.org/package/2006/relationships",
  contentTypes: "http://schemas.openxmlformats.org/package/2006/content-types",
  app: "http://schemas.microsoft.com/sharepoint/2012/app/manifest",
  sharePoint: "http://schemas.microsoft.com/sharepoint/",
  partConfig: "http://schemas.microsoft.com/sharepoint/2012/app/partconfiguration",
} as const;

const REL_TYPE = "http://schemas.microsoft.com/sharepoint/2012/app/relationships";

/** Server-side contract string — exact casing and trailing slash matter. */
const ASSET_LIBRARY = "HTTPS://SPCLIENTSIDEASSETLIBRARY/";

const CONTENT_TYPES: Record<string, string> = {
  xml: "text/xml",
  rels: "application/vnd.openxmlformats-package.relationships+xml",
  webpart: "text/xml",
  htm: "text/html",
  html: "text/html",
  aspx: "text/xml",
  resx: "text/xml",
  js: "application/javascript",
  json: "application/json",
  png: "image/png",
  jpg: "image/jpeg",
  bmp: "image/bmp",
  gif: "image/gif",
};

// ---------------------------------------------------------------------- xml

const DECLARATION = '<?xml version="1.0" encoding="utf-8"?>';

const escapeAttribute = (value: unknown): string =>
  String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");

const escapeText = (value: unknown): string =>
  String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** Serialise attributes, skipping undefined values. */
const attributes = (values: Record<string, string | undefined>): string =>
  Object.entries(values)
    .filter((entry): entry is [string, string] => entry[1] !== undefined)
    .map(([key, value]) => ` ${key}="${escapeAttribute(value)}"`)
    .join("");

interface Relationship {
  type: string;
  target: string;
}

function relationshipsXml(relationships: Relationship[]): string {
  const items = relationships
    .map(
      ({ type, target }, index) =>
        `<Relationship${attributes({
          Type: `${REL_TYPE}/${type}`,
          Target: `/${target}`,
          Id: `r${index}`,
        })}></Relationship>`,
    )
    .join("");
  return `${DECLARATION}<Relationships${attributes({ xmlns: NS.relationships })}>${items}</Relationships>`;
}

/**
 * Stable UUIDv5-shaped GUID. Heft randomises these on every build; deriving them
 * makes packages byte-reproducible, and SharePoint has no stake in the value.
 */
function stableGuid(seed: string): string {
  const hex = createHash("sha1").update(seed).digest("hex").slice(0, 32).split("");
  hex[12] = "5"; // version
  hex[16] = "89ab"[parseInt(hex[16]!, 16) % 4]!; // RFC 4122 variant
  const s = hex.join("");
  return `${s.slice(0, 8)}-${s.slice(8, 12)}-${s.slice(12, 16)}-${s.slice(16, 20)}-${s.slice(20, 32)}`;
}

const appPartConfigXml = (id: string): string =>
  `${DECLARATION}<AppPartConfig${attributes({ xmlns: NS.partConfig })}><Id>${escapeText(id)}</Id></AppPartConfig>`;

// --------------------------------------------------------------- generators

function appManifestXml(solution: Solution): string {
  const developer = solution.developer ?? {};
  const metadata = solution.metadata ?? {};

  const localized = (dictionary: Record<string, string> | undefined, tag: string): string => {
    if (!dictionary) return "";
    const inner = Object.entries(dictionary)
      .map(
        ([culture, text]) =>
          `<LocalizedString${attributes({ CultureName: culture })}>${escapeText(text)}</LocalizedString>`,
      )
      .join("");
    return `<${tag}>${inner}</${tag}>`;
  };

  const properties =
    `<Title>${escapeText(solution.title ?? solution.name)}</Title>` +
    `<DeveloperProperties>${escapeText(
      JSON.stringify({
        name: developer.name ?? "",
        websiteUrl: developer.websiteUrl ?? "",
        privacyUrl: developer.privacyUrl ?? "",
        termsOfUseUrl: developer.termsOfUseUrl ?? "",
        mpnId: developer.mpnId ?? "",
      }),
    )}</DeveloperProperties>` +
    localized(metadata.shortDescription, "ShortDescription") +
    localized(metadata.longDescription, "LongDescription") +
    (metadata.categories?.length
      ? `<CategoryID>${escapeText(metadata.categories.slice(0, 3).join(","))}</CategoryID>`
      : "") +
    (metadata.videoUrl ? `<VideoUrl>${escapeText(metadata.videoUrl)}</VideoUrl>` : "");

  const permissions = solution.webApiPermissionRequests?.length
    ? `<WebApiPermissionRequests>${solution.webApiPermissionRequests
        .map(
          (request) =>
            `<WebApiPermissionRequest${attributes({
              ResourceId: request.resource,
              Scope: request.scope,
              AppId: request.appId,
              ReplyUrl: request.replyUrl,
            })}></WebApiPermissionRequest>`,
        )
        .join("")}</WebApiPermissionRequests>`
    : "";

  return `${DECLARATION}<App${attributes({
    xmlns: NS.app,
    Name: solution.name,
    ProductID: solution.id,
    SharePointMinVersion: "16.0.0.0",
    IsClientSideSolution: "true",
    Version: solution.version ?? "1.0.0.0",
    SkipFeatureDeployment: solution.skipFeatureDeployment ? "true" : undefined,
    IsDomainIsolated:
      solution.isDomainIsolated === undefined ? undefined : String(!!solution.isDomainIsolated),
  })}><Properties>${properties}</Properties>${permissions}</App>`;
}

const featureXml = (feature: {
  title: string;
  description: string;
  id: string;
  version?: string;
}): string =>
  `${DECLARATION}<Feature${attributes({
    xmlns: NS.sharePoint,
    Title: feature.title,
    Description: feature.description,
    Id: feature.id,
    Version: feature.version ?? "1.0.0.0",
    Scope: "Web",
    Hidden: "FALSE",
  })}></Feature>`;

/**
 * The element manifest is where the component manifest JSON actually lives.
 * `<Module Url="_catalogs/wp" List="113">` provisions the web part gallery entry
 * and applies only to WebPart / AdaptiveCardExtension components.
 */
function componentElementXml(manifest: ComponentManifest, displayName: string): string {
  const needsModule =
    manifest.componentType === "WebPart" || manifest.componentType === "AdaptiveCardExtension";

  const packed: ComponentManifest = { ...manifest };
  delete packed.requiresCustomScript;
  delete packed.safeWithCustomScriptDisabled;

  return (
    `${DECLARATION}<Elements${attributes({ xmlns: NS.sharePoint })}>` +
    `<ClientSideComponent${attributes({
      Name: displayName,
      Id: manifest.id,
      ComponentManifest: JSON.stringify(packed),
      Type: manifest.componentType,
      ReturnIfCustomScriptDisabled: manifest.requiresCustomScript ? "false" : undefined,
    })}></ClientSideComponent>` +
    (needsModule
      ? `<Module${attributes({ Name: displayName, Url: "_catalogs/wp", List: "113" })}></Module>`
      : "") +
    `</Elements>`
  );
}

const contentTypesXml = (extensions: string[]): string =>
  `${DECLARATION}<Types${attributes({ xmlns: NS.contentTypes })}>${extensions
    .map(
      (ext) =>
        `<Default${attributes({
          Extension: ext,
          ContentType: CONTENT_TYPES[ext] ?? "application/octet-stream",
        })}></Default>`,
    )
    .join("")}</Types>`;

/** How SPFx derives the name shown in the toolbox and the package. */
export const displayNameFor = (manifest: ComponentManifest): string =>
  (manifest.componentType === "WebPart"
    ? manifest.preconfiguredEntries?.[0]?.title?.default
    : undefined) ?? manifest.alias;

// -------------------------------------------------------------------- packer

export interface SppkgInput {
  solution: Solution;
  /** Release component manifests, `loaderConfig` included. */
  manifests: ComponentManifest[];
  /** Files destined for `ClientSideAssets/`, with their emitted (hashed) names. */
  assets: PackagedAsset[];
}

export function createSppkg({ solution, manifests, assets }: SppkgInput): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  const put = (name: string, content: string | Uint8Array): void => {
    files[name] = typeof content === "string" ? Buffer.from(content, "utf8") : content;
  };

  const embedsAssets = solution.includeClientSideAssets !== false;

  const packed = manifests.map((manifest) => {
    const copy = structuredClone(manifest);
    if (embedsAssets && copy.loaderConfig) {
      copy.loaderConfig.internalModuleBaseUrls = [ASSET_LIBRARY];
    }
    return copy;
  });

  // Without an explicit feature list SPFx generates one feature per component.
  const declared: SolutionFeature[] = solution.features?.length
    ? solution.features
    : packed.map((manifest) => ({
        id: manifest.id,
        title: `${displayNameFor(manifest)} Feature`,
        description: `The feature that activates ${displayNameFor(manifest)}.`,
        version: "1.0.0.0",
      }));

  const appManifestRelationships: Relationship[] = [];

  declared.forEach((feature, index) => {
    // A feature without componentIds claims everything not yet assigned.
    const components = feature.componentIds
      ? packed.filter((manifest) => feature.componentIds!.includes(manifest.id))
      : index === 0
        ? packed
        : [];

    const file = `feature_${feature.id}.xml`;
    put(file, featureXml(feature));
    put(`${file}.config.xml`, appPartConfigXml(stableGuid(`${solution.id}:${feature.id}`)));

    const relationships: Relationship[] = [
      { type: "partconfiguration", target: `${file}.config.xml` },
    ];
    for (const manifest of components) {
      const elementFile = `${feature.id}/${manifest.componentType}_${manifest.id}.xml`;
      put(elementFile, componentElementXml(manifest, displayNameFor(manifest)));
      relationships.push({ type: "feature-elementmanifest", target: elementFile });
    }
    put(`_rels/${file}.rels`, relationshipsXml(relationships));
    appManifestRelationships.push({ type: "manifest-feature", target: file });
  });

  const clientSideAssets = embedsAssets ? assets : [];
  if (clientSideAssets.length > 0) {
    const file = "ClientSideAssets.xml";
    put(
      file,
      featureXml({
        title: "Client Side Assets",
        description:
          "A feature that help deploy client side component assets to SharePoint Online.",
        id: stableGuid(`${solution.id}:clientsideassets`),
        version: "1.0.0.0",
      }),
    );
    put(
      `${file}.config.xml`,
      appPartConfigXml(stableGuid(`${solution.id}:clientsideassets:config`)),
    );

    const relationships: Relationship[] = [
      { type: "partconfiguration", target: `${file}.config.xml` },
    ];
    for (const asset of clientSideAssets) {
      put(`ClientSideAssets/${asset.name}`, asset.data);
      relationships.push({ type: "clientsideasset", target: `ClientSideAssets/${asset.name}` });
    }
    put(`_rels/${file}.rels`, relationshipsXml(relationships));
    appManifestRelationships.push({ type: "manifest-clientsideasset", target: file });
  }

  put("AppManifest.xml", appManifestXml(solution));
  put("_rels/AppManifest.xml.rels", relationshipsXml(appManifestRelationships));
  put("_rels/.rels", relationshipsXml([{ type: "package-manifest", target: "AppManifest.xml" }]));

  // OPC requires a content type for every extension present in the package.
  const extensions = new Set(["xml", "rels"]);
  for (const name of Object.keys(files)) {
    const ext = name.includes(".") ? name.slice(name.lastIndexOf(".") + 1).toLowerCase() : "";
    if (ext) extensions.add(ext);
  }
  put("[Content_Types].xml", contentTypesXml([...extensions].toSorted()));

  // Fixed mtime keeps rebuilds byte-identical.
  return zipSync(files, { level: 9, mtime: new Date("2000-01-01T00:00:00Z") });
}
