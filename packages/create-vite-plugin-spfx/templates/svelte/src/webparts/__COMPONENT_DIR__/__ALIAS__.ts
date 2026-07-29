import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import { Version } from "@microsoft/sp-core-library";
import {
  PropertyPaneTextField,
  PropertyPaneToggle,
  type IPropertyPaneConfiguration,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { applyTheme } from "vite-plugin-spfx/runtime";
import { mount, unmount } from "svelte";

import * as strings from "__STRINGS__";
import __COMPONENT__ from "./__COMPONENT__.svelte";
import { props } from "./props.svelte.ts";

export interface I__ALIAS__Props {
  description: string;
  showDetails: boolean;
}

export default class __ALIAS__ extends BaseClientSideWebPart<I__ALIAS__Props> {
  private _component: Record<string, unknown> | undefined;

  /**
   * SPFx calls render() again on every property pane change. Remounting would
   * reset component state, so mount once and mutate the shared `$state` — the
   * Svelte equivalent of what React's reconciliation does for free.
   */
  public render(): void {
    Object.assign(props, this._current());
    this._component ??= mount(__COMPONENT__, { target: this.domElement });
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    applyTheme(this.domElement, theme);
  }

  protected onDispose(): void {
    if (this._component) unmount(this._component);
    this._component = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyPaneToggle("showDetails", {
                  label: strings.ShowDetailsFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  /** Everything the component needs, resolved once — identical in all examples. */
  private _current() {
    return {
      description: this.properties.description,
      showDetails: this.properties.showDetails,
      userDisplayName: this.context.pageContext.user.displayName,
      environment: this.context.isServedFromLocalhost
        ? strings.LocalEnvironment
        : strings.SharePointEnvironment,
      labels: {
        greeting: strings.Greeting,
        property: strings.PropertyLabel,
        clickCount: strings.ClickCount,
      },
    };
  }
}
