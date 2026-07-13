// @ts-check
import { defineConfig } from "astro/config";
import svelte from '@astrojs/svelte';
import UnoCSS from "unocss/astro";


// https://astro.build/config
export default defineConfig({
  site: "https://tdjw.chenxing.dev",
  base: "/",
  integrations: [
    svelte(),
    UnoCSS({
      injectReset: true,
    }),
  ],
});
