<script lang="ts">
  import { onMount } from "svelte";
  import { 匹配所属社区 } from "@core/resolver";

  let input = $state("");
  let addrInputEl: HTMLInputElement | null = null;

  const output = $derived(input ? 匹配所属社区(input) || "？？？" : "？？？");

  function clearInput() {
    input = "";
    addrInputEl?.focus();
  }

  onMount(() => {
    const isDesktop = window.matchMedia("(min-width: 768px)").matches;
    const isBodyFocused = document.activeElement === document.body;

    if (isDesktop && isBodyFocused) {
      requestAnimationFrame(() => {
        addrInputEl?.focus();
      });
    }

    return () => {};
  });
</script>

<div
  class="grid grid-cols-1 md:grid-cols-[auto_1fr] border border-black mt-4 w-full"
>
  <div class="md:border-r border-black px-2 py-1">
    <label for="addrInput" class="font-bold whitespace-nowrap">请输入地址</label
    >
  </div>
  <div class="border-t md:border-none px-2 py-1">
    <div class="relative">
      <input
        id="addrInput"
        name="address"
        type="text"
        inputmode="text"
        placeholder="例如：西槎路31号"
        class="w-full bg-white pr-12 outline-none placeholder:text-black/45 transition border focus-visible:border-black focus-visible:ring-2 focus-visible:ring-black/20 border-black ring-2 ring-black/20"
        bind:this={addrInputEl}
        bind:value={input}
      />
      {#if input}
        <button
          type="button"
          class="absolute right-2 top-1/2 -translate-y-1/2 text-sm text-black/70 hover:text-black"
          onclick={clearInput}
          aria-label="清空输入"
        >
          清空
        </button>
      {/if}
    </div>
  </div>
  <div class="border-t md:border-r border-black px-2 py-1">
    <span class="font-bold whitespace-nowrap">所属社区</span>
  </div>
  <div class="border-t border-black px-2 py-1">{output}</div>
</div>
