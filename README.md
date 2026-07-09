# 自动化分居委

自动化分居委是一个用 JavaScript 写的简易小工具，目的在于将给定的某街道辖区内的任一地址对应到其所属社区（aka 居委）。

## 原理
脚本会读取一个.xlsx，按你指定的“地址列”匹配居委并把结果写回表格。

## 具体怎么操作

### 安装依赖

1. 安装 [Node.js v11.9.0+](https://nodejs.org/en/download).
2. 克隆本仓库并跑 `npm install` 安装依赖:
```bash
git clone https://github.com/dev-chenxing/neighborhood-locator.git
cd neighborhood-locator
npm install
```

### 运行脚本

在仓库根目录用命令行运行以下命令：

```bash
node ./neighborhood_locator.js "\path\to\你的名单.xlsx" "地址" "所在居委" --log-level ERROR
```

说明：
- 第一个参数：输入文件路径（仅支持.xlsx文件）。
- 第二个参数：地址列名，如“地址”。
- 第三个参数：输出列名，如“所在居委”（脚本会把匹配结果填到该列）。

### 作为模块使用

如果你要在别的项目里直接复用解析逻辑，可以直接从包入口导入：

```js
const { 匹配所属社区 } = require("neighborhood-locator");

const 社区 = 匹配所属社区("xx路129号");
```

#### Shell Script 全自動 (不推荐)

`分居委` 腳本是針對某特定格式的辦公表格而寫的全自動腳本，只需一行命令即可完成 99%分居委的工作，但需要安裝`Shell`, 如：[Git Bash](https://git-scm.com/downloads)

```bash
./分居委 2024年5月区新开办企业情况清单.xlsx <區> <街道>
```
