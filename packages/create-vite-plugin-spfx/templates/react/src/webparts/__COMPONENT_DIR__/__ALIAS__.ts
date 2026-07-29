import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import { Version } from "@microsoft/sp-core-library";
import {
  PropertyPaneTextField,
  PropertyPaneToggle,
  type IPropertyPaneConfiguration,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { applyTheme } from "vite-plugin-spfx/runtime";
import * as React from "react";
import * as ReactDom from "react-dom";

import * as strings from "__STRINGS__";
import __COMPONENT__ from "./__COMPONENT__.tsx";

export interface I__ALIAS__Props {
  description: string;
  showDetails: boolean;
}

export default class __ALIAS__ extends BaseClientSideWebPart<I__ALIAS__Props> {
  public render(): void {
    // ReactDom.render reconciles into the same container, so component state
    // survives the re-render SPFx triggers on every property pane change.
    ReactDom.render(React.createElement(__COMPONENT__, this._props()), this.domElement);
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    applyTheme(this.domElement, theme);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
  private _props() {
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
