# 自动化分居委

自动化分居委是一个用 Python 和 JavaScript 写的简易小工具，目的在于将给定的某街道辖区内的任一地址对应到其所属社区（aka 居委）

## 如何使用Python版（推荐）

### 前期准备

安装`Python 3.9+`. 如果`Python 3.9+`不支持你的电脑系统，请使用[JavaScript版](#如何使用javascript版)

### 安装

克隆这个仓库到你的电脑:

```bash
git clone https://github.com/amaliegay/neighborhood-locator.git
```

cd 到克隆目录下面后，创建虚拟环境:

```powershell
cd neighborhood-locator
python -m venv venv
venv/Scripts/activate # source venv/bin/activate 如果是Linux系统
```

使用 pip 安装所需的 Python 包:

```bash
pip install -r requirements.txt
```

### 最后一步

在命令行运行 Python 脚本:

```bash
python locator.py <input.xlsx> <col_name>
```

这里，`<input.xlsx>`是输入地址所在的 Excel 表文件名，`<col_name>`是表中地址一列的表头名。
例如，如果输入的 Excel 表是`新开办企业情况清单.xlsx`, `住所`是地址一列的表头名，则应运行:

```bash
python locator.py 新开办企业情况清单.xlsx 住所
```

脚本运行完毕时，命令行会输出地址和所属居委的列表，同时输入的Excel表也会更新所属居委列表。

## 如何使用JavaScript版

### 前期准备

安装 [Node.js v11.9.0+](https://nodejs.org/en/download).

### 安装

克隆这个仓库到你的电脑:

```bash
git clone https://github.com/amaliegay/neighborhood-locator.git
```

cd 到克隆目录下面后，使用 npm 安装所需的包:

```powershell
cd neighborhood-locator
npm install
```

### 最后一步

在命令行运行 JavaScript 脚本:

```bash
node index.js <input.xlsx> <col_name>
```

这里，`<input.xlsx>`是输入地址所在的 Excel 表文件名，`<col_name>`是表中地址一列的表头名。
例如，如果输入的 Excel 表是`新开办企业情况清单.xlsx`, `住所`是地址一列的表头名，则应运行:

```bash
node index.js 新开办企业情况清单.xlsx 住所
```

脚本运行完毕时，命令行会输出地址和所属居委的列表，同时输入的Excel表也会更新所属居委列表。
