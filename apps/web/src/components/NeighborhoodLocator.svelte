<script lang="ts">
  import { 匹配所属社区 } from "@core/resolver";

  let input = $state("");

  const output = $derived(input ? 匹配所属社区(input) || "？？？" : "？？？");
</script>

<div class="grid grid-cols-1 md:grid-cols-[auto_1fr] border border-black mt-4 w-full">
  <div class="md:border-r border-black px-2 py-1">
    <label for="addrInput" class="font-bold whitespace-nowrap">地址</label>
  </div>
  <div class="border-t md:border-none px-2 py-1">
    <input
      type="text"
      placeholder="输入地址，例如：西槎路31号"
      class="w-full outline-none"
      bind:value={input}
    />
  </div>
  <div class="border-t md:border-r border-black px-2 py-1">
    <span class="font-bold whitespace-nowrap">所属社区</span>
  </div>
  <div class="border-t border-black px-2 py-1">{output}</div>
</div>
