<script lang="ts">
  import { 匹配所属社区 } from "@core/resolver";

  let input = $state("");

  const output = $derived(input ? 匹配所属社区(input) || "？？？" : "");
</script>

<!-- 移动端上下，桌面端左右 -->
<div class="flex flex-col md:flex-row items-center gap-8">
  <div class="md:w-80 shrink-0">
    <input
      type="text"
      id="addrInput"
      placeholder="输入地址，例如：西槎路31号"
      class="w-full border px-3 py-2 rounded"
      bind:value={input}
    />
  </div>
  <div class="grow">
    <div class="text-xl">{output}</div>
  </div>
</div>
