import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import { Version } from "@microsoft/sp-core-library";
import {
  PropertyPaneTextField,
  PropertyPaneToggle,
  type IPropertyPaneConfiguration,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { applyTheme } from "vite-plugin-spfx/runtime";

import * as strings from "__STRINGS__";
import { create__COMPONENT__, type I__COMPONENT__Component } from "./__COMPONENT__.ts";

export interface I__ALIAS__Props {
  description: string;
  showDetails: boolean;
}

export default class __ALIAS__ extends BaseClientSideWebPart<I__ALIAS__Props> {
  private _component: I__COMPONENT__Component | undefined;

  /**
   * SPFx calls render() again on every property pane change. Rebuilding the DOM
   * each time would reset component state, so create it once and update it —
   * the same split the framework examples get from their renderer, only visible.
   */
  public render(): void {
    if (!this._component) {
      this._component = create__COMPONENT__();
      this.domElement.append(this._component.element);
    }
    this._component.update(this._current());
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    applyTheme(this.domElement, theme);
  }

  protected onDispose(): void {
    this._component?.element.remove();
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
