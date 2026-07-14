<script lang="ts">
  import { 匹配所属社区 } from "@core/resolver";

  let input = $state("");

  const output = $derived(input ? 匹配所属社区(input) || "？？？" : "？？？");
</script>

<div class="grid grid-cols-[auto_1fr] w-fit border border-black mt-4 w-full">
  <div class="border-r border-black px-2">
    <label for="addrInput" class="font-bold whitespace-nowrap">地址</label>
  </div>
  <div class="min-w-[22rem] px-2">
    <input
      type="text"
      placeholder="输入地址，例如：西槎路31号"
      class="w-full outline-none"
      bind:value={input}
    />
  </div>
  <div class="border-r border-black border-t px-2">
    <span class="font-bold whitespace-nowrap">所属社区</span>
  </div>
  <div class="border-t border-black px-2">{output}</div>
</div>
