declare interface I__STRINGS__ {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ShowDetailsFieldLabel: string;
  Greeting: string;
  PropertyLabel: string;
  ClickCount: string;
  LocalEnvironment: string;
  SharePointEnvironment: string;
}

/**
 * The bare specifier is resolved by the plugin's `localizedResources` option and
 * emitted as one standalone AMD module per locale. Nothing generates this file —
 * it is maintained by hand, exactly as in a stock SPFx project.
 */
declare module "__STRINGS__" {
  const strings: I__STRINGS__;
  export = strings;
}
