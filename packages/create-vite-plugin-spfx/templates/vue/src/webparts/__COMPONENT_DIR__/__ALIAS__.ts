import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import { Version } from "@microsoft/sp-core-library";
import {
  PropertyPaneTextField,
  PropertyPaneToggle,
  type IPropertyPaneConfiguration,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { applyTheme } from "vite-plugin-spfx/runtime";
import { createApp, h, reactive, type App } from "vue";

import * as strings from "__STRINGS__";
import __COMPONENT__ from "./__COMPONENT__.vue";

export interface I__ALIAS__Props {
  description: string;
  showDetails: boolean;
}

export default class __ALIAS__ extends BaseClientSideWebPart<I__ALIAS__Props> {
  private _app: App | undefined;

  /**
   * SPFx calls render() again on every property pane change. Recreating the app
   * each time would reset component state and rebuild the DOM, so mount once and
   * only update this reactive object — React's ReactDom.render() does the
   * equivalent implicitly.
   */
  private readonly _props = reactive({
    description: "",
    showDetails: false,
    userDisplayName: "",
    environment: "",
    labels: { greeting: "", property: "", clickCount: "" },
  });

  public render(): void {
    Object.assign(this._props, this._current());

    if (this._app) return;

    const props = this._props;
    // Spreading inside the render function is what makes each property a tracked
    // dependency of the root render effect.
    this._app = createApp({ render: () => h(__COMPONENT__, { ...props }) });
    this._app.mount(this.domElement);
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    applyTheme(this.domElement, theme);
  }

  protected onDispose(): void {
    this._app?.unmount();
    this._app = undefined;
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
