/**
 * Reactive props shared between the web part and the Svelte component.
 *
 * Runes only work in `.svelte` and `.svelte.ts` files, which is why this lives
 * apart from the web part class. Mutating this object updates the mounted
 * component in place — the Svelte equivalent of React's reconciliation.
 */
export const props = $state({
  description: "",
  showDetails: false,
  userDisplayName: "",
  environment: "",
  labels: { greeting: "", property: "", clickCount: "" },
});
