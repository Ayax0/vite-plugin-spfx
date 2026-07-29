/** Options accepted by the {@link spfx} plugin. */
export interface SpfxOptions {
  /**
   * Bundle key, as in SPFx's `config.json`. Becomes `loaderConfig.entryModuleId`
   * and the emitted file name.
   */
  bundleName: string;
  /** Component entry, relative to the project root. */
  entry: string;
  /** The component's `*.manifest.json`, relative to the project root. */
  manifest: string;
  /**
   * Bare import specifier -> per-locale files, e.g.
   * `{ MyWebPartStrings: 'src/webparts/my/loc/{locale}.js' }`.
   * Each locale is emitted as a standalone AMD module and referenced from the
   * component manifest.
   */
  localizedResources?: Record<string, string>;
  /**
   * Base URL written into the debug manifest. Defaults to the dev server's own
   * origin, derived from `config/serve.json` or the resolved port.
   */
  serveBaseUrl?: string;
  /** Base URL written into the release manifest; empty ships assets in the package. */
  cdnBasePath?: string;
  /**
   * Page the dev server prints a ready-made debug link for, e.g.
   * `https://{tenantDomain}/_layouts/workbench.aspx`. Falls back to
   * `config/serve.json`; `{tenantDomain}` comes from SPFX_SERVE_TENANT_DOMAIN.
   */
  serveInitialPage?: string;
}

/** How the SharePoint loader resolves one entry of `loaderConfig.scriptResources`. */
export type ScriptResource =
  | { type: "path"; path: string }
  | { type: "component"; id: string; version: string }
  | { type: "localizedPath"; defaultPath: string; paths: Record<string, string> };

export interface LoaderConfig {
  /** Ordered failover base URLs; every file must exist on every host. */
  internalModuleBaseUrls: string[];
  /** Names the `scriptResources` entry holding the component itself. */
  entryModuleId: string;
  scriptResources: Record<string, ScriptResource>;
}

export type ComponentType =
  | "WebPart"
  | "Extension"
  | "Library"
  | "Application"
  | "AdaptiveCardExtension";

export interface ComponentManifest {
  id: string;
  alias: string;
  componentType: ComponentType;
  /** `"*"` in source, replaced with the package version at build time. */
  version: string;
  manifestVersion: number;
  requiresCustomScript?: boolean;
  safeWithCustomScriptDisabled?: boolean;
  preconfiguredEntries?: Array<{ title?: { default?: string } } & Record<string, unknown>>;
  loaderConfig?: LoaderConfig;
  [key: string]: unknown;
}

/** A package SharePoint already serves at runtime. */
export interface SpfxComponent {
  id: string;
  version: string;
}

export interface SolutionFeature {
  id: string;
  title: string;
  description: string;
  version?: string;
  /** Restricts the feature to these components; otherwise it takes all unclaimed ones. */
  componentIds?: string[];
}

/** The `solution` block of `config/package-solution.json`. */
export interface Solution {
  name: string;
  id: string;
  version?: string;
  title?: string;
  includeClientSideAssets?: boolean;
  skipFeatureDeployment?: boolean;
  isDomainIsolated?: boolean;
  developer?: {
    name?: string;
    websiteUrl?: string;
    privacyUrl?: string;
    termsOfUseUrl?: string;
    mpnId?: string;
  };
  metadata?: {
    shortDescription?: Record<string, string>;
    longDescription?: Record<string, string>;
    categories?: string[];
    videoUrl?: string;
  };
  features?: SolutionFeature[];
  webApiPermissionRequests?: Array<{
    resource: string;
    scope: string;
    appId?: string;
    replyUrl?: string;
  }>;
}

export interface PackageSolutionConfig {
  solution: Solution;
  paths?: { zippedPackage?: string };
}

/** A file destined for `ClientSideAssets/` in the package. */
export interface PackagedAsset {
  name: string;
  data: Uint8Array;
}
