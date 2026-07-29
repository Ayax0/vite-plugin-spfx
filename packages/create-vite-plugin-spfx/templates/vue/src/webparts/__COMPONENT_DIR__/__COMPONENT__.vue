<script setup lang="ts">
import { ref } from "vue";

import logo from "./assets/logo.svg";

const props = defineProps<{
  description: string;
  showDetails: boolean;
  userDisplayName: string;
  environment: string;
  labels: { greeting: string; property: string; clickCount: string };
}>();

const count = ref(0);
</script>

<template>
  <section class="hello">
    <!-- The asset URL is resolved at runtime from the script's own location. -->
    <img class="logo" :src="logo" alt="" />
    <h2>{{ props.labels.greeting }} {{ props.userDisplayName }} — this is Vue</h2>
    <p>
      {{ props.labels.property }}: <strong>{{ props.description }}</strong>
    </p>
    <p v-if="props.showDetails" class="details">{{ props.environment }}</p>
    <!-- The count survives property pane edits — proof that the web part patches
         props rather than remounting. -->
    <button type="button" @click="count++">
      {{ props.labels.clickCount.replace("{0}", String(count)) }}
    </button>
  </section>
</template>

<style scoped lang="scss">
.hello {
  padding: 1em;
  // The plugin rewrites this to var(--bodyText, #323130); applyTheme publishes
  // the slot from the host theme.
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
