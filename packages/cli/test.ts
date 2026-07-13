import { 匹配所属社区 } from "../core/resolver";

const 输入地址 = process.argv.slice(2).join(" ");

if (!输入地址) {
  console.error("❌ 请输入要查询的地址，例如：bun test 荔德路318号");
  process.exit(1);
}

const 社区名称 = 匹配所属社区(输入地址);

if (社区名称) {
  console.log(社区名称);
} else {
  console.log(`❌ 未找到地址 "${输入地址}" 的所属社区`);
}
