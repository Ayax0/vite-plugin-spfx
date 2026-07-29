import { createStore } from "solid-js/store";

export interface I__COMPONENT__Props {
  description: string;
  showDetails: boolean;
  userDisplayName: string;
  environment: string;
  labels: { greeting: string; property: string; clickCount: string };
}

/**
 * Reactive props shared between the web part and the Solid component.
 *
 * A store rather than signals: the web part hands over a whole object on every
 * render(), and `setProps(object)` merges it field by field, so only what
 * actually changed re-runs. Reading `props.x` inside the component is what
 * subscribes to it.
 */
export const [props, setProps] = createStore<I__COMPONENT__Props>({
  description: "",
  showDetails: false,
  userDisplayName: "",
  environment: "",
  labels: { greeting: "", property: "", clickCount: "" },
});
