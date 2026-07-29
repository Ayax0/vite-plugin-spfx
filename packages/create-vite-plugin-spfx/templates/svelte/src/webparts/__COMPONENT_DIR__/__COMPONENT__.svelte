<script lang="ts">
  import logo from "./assets/logo.svg";
  import { props } from "./props.svelte.ts";

  let count = $state(0);
</script>

<section class="hello">
  <!-- The asset URL is resolved at runtime from the script's own location. -->
  <img class="logo" src={logo} alt="" />
  <h2>{props.labels.greeting} {props.userDisplayName} — this is Svelte</h2>
  <p>{props.labels.property}: <strong>{props.description}</strong></p>
  {#if props.showDetails}
    <p class="details">{props.environment}</p>
  {/if}
  <!-- The count survives property pane edits: the web part mutates the shared
       $state instead of remounting the component. -->
  <button type="button" onclick={() => count++}>
    {props.labels.clickCount.replace("{0}", String(count))}
  </button>
</section>

<style lang="scss">
  .hello {
    padding: 1em;
    // The plugin rewrites this to var(--bodyText, #323130); applyTheme
    // publishes the slot from the host theme.
    color: "[theme:bodyText, default: #323130]";
  }

  .logo {
    width: 48px;
    height: 48px;
  }

  .details {
    font-style: italic;
    opacity: 0.8;
  }

  button {
    padding: 0.4em 1em;
    border: 1px solid var(--link);
    border-radius: 4px;
    background: transparent;
    cursor: pointer;
    color: "[theme:link, default: #03787c]";
  }
</style>
