import logo from "./assets/logo.svg";
import styles from "./styles.module.scss";

export interface I__COMPONENT__Props {
  description: string;
  showDetails: boolean;
  userDisplayName: string;
  environment: string;
  labels: { greeting: string; property: string; clickCount: string };
}

export interface I__COMPONENT__Component {
  element: HTMLElement;
  update(props: I__COMPONENT__Props): void;
}

/**
 * Build the DOM once and hand back a way to update it.
 *
 * The other examples get this shape from their framework. Written out by hand it
 * is the whole trick: nodes are created in the factory, and `update` only writes
 * to them — so nothing that lives in a node (focus, scroll position, the click
 * count below) is destroyed when SPFx calls render() again.
 */
export function create__COMPONENT__(): I__COMPONENT__Component {
  let count = 0;
  let clickCountLabel = "";

  const section = document.createElement("section");
  section.className = styles.hello;

  // The asset URL is resolved at runtime from the script's own location.
  const image = document.createElement("img");
  image.className = styles.logo;
  image.src = logo;
  image.alt = "";

  const heading = document.createElement("h2");

  const property = document.createElement("p");
  const propertyLabel = document.createTextNode("");
  const propertyValue = document.createElement("strong");
  property.append(propertyLabel, propertyValue);

  const details = document.createElement("p");
  details.className = styles.details;

  const button = document.createElement("button");
  button.type = "button";
  button.className = styles.button;

  const paintButton = (): void => {
    button.textContent = clickCountLabel.replace("{0}", String(count));
  };

  // The count survives property pane edits — it lives in this closure, which
  // update() never touches.
  button.addEventListener("click", () => {
    count += 1;
    paintButton();
  });

  section.append(image, heading, property, details, button);

  return {
    element: section,
    update(props: I__COMPONENT__Props): void {
      heading.textContent =
        `${props.labels.greeting} ${props.userDisplayName} — this is plain TypeScript`;
      propertyLabel.textContent = `${props.labels.property}: `;
      propertyValue.textContent = props.description;
      details.textContent = props.environment;
      details.hidden = !props.showDetails;
      clickCountLabel = props.labels.clickCount;
      paintButton();
    },
  };
}
