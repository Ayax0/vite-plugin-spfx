import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import { Version } from "@microsoft/sp-core-library";
import {
  PropertyPaneTextField,
  PropertyPaneToggle,
  type IPropertyPaneConfiguration,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { applyTheme } from "vite-plugin-spfx/runtime";
import { createComponent, render } from "solid-js/web";

import * as strings from "__STRINGS__";
import __COMPONENT__ from "./__COMPONENT__.tsx";
import { props, setProps } from "./props.ts";

export interface I__ALIAS__Props {
  description: string;
  showDetails: boolean;
}

export default class __ALIAS__ extends BaseClientSideWebPart<I__ALIAS__Props> {
  private _dispose: (() => void) | undefined;

  /**
   * SPFx calls render() again on every property pane change. Rendering a second
   * time would build a second tree, so mount once and update the store — the
   * Solid equivalent of what React's reconciliation does for free.
   */
  public render(): void {
    setProps(this._current());

    // createComponent is what JSX compiles to. Using it directly keeps this file
    // plain TypeScript and passes the store itself, so the component's own reads
    // are what establish the fine-grained subscriptions.
    this._dispose ??= render(() => createComponent(__COMPONENT__, props), this.domElement);
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    applyTheme(this.domElement, theme);
  }

  protected onDispose(): void {
    this._dispose?.();
    this._dispose = undefined;
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
