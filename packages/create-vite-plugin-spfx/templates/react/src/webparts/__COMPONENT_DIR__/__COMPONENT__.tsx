import * as React from "react";

import logo from "./assets/logo.svg";
import styles from "./styles.module.scss";

export interface I__COMPONENT__Props {
  description: string;
  showDetails: boolean;
  userDisplayName: string;
  environment: string;
  labels: { greeting: string; property: string; clickCount: string };
}

export default function __COMPONENT__(
  props: I__COMPONENT__Props,
): React.ReactElement {
  const [count, setCount] = React.useState(0);

  return (
    <section className={styles.hello}>
      {/* The asset URL is resolved at runtime from the script's own location. */}
      <img className={styles.logo} src={logo} alt="" />
      <h2>
        {props.labels.greeting} {props.userDisplayName} — this is React
      </h2>
      <p>
        {props.labels.property}: <strong>{props.description}</strong>
      </p>
      {props.showDetails && (
        <p className={styles.details}>{props.environment}</p>
      )}
      {/* The count survives property pane edits — proof that render() patches
          rather than remounts. */}
      <button
        type="button"
        className={styles.button}
        onClick={() => setCount(count + 1)}
      >
        {props.labels.clickCount.replace("{0}", String(count))}
      </button>
    </section>
  );
}
