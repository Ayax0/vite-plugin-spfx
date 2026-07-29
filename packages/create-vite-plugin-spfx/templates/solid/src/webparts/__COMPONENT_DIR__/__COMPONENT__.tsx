import { createSignal, Show, type JSX } from "solid-js";

import logo from "./assets/logo.svg";
import styles from "./styles.module.scss";
import type { I__COMPONENT__Props } from "./props.ts";

export default function __COMPONENT__(props: I__COMPONENT__Props): JSX.Element {
  const [count, setCount] = createSignal(0);

  return (
    <section class={styles.hello}>
      {/* The asset URL is resolved at runtime from the script's own location. */}
      <img class={styles.logo} src={logo} alt="" />
      <h2>
        {props.labels.greeting} {props.userDisplayName} — this is Solid
      </h2>
      <p>
        {props.labels.property}: <strong>{props.description}</strong>
      </p>
      <Show when={props.showDetails}>
        <p class={styles.details}>{props.environment}</p>
      </Show>
      {/* The count survives property pane edits — proof that the web part
          updates the store rather than remounting. */}
      <button type="button" class={styles.button} onClick={() => setCount(count() + 1)}>
        {props.labels.clickCount.replace("{0}", String(count()))}
      </button>
    </section>
  );
}
