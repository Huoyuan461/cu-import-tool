# CU导入工具

![Tenneco Logo](docs/assets/tenneco-logo.png)

CU导入工具是一个本机可运行的 Excel 写入工具，用于把源预测文件中的产品数据，按规则写入目标容量模板。它支持：

- 浏览器 HTML 版，无需安装 Python
- Windows HTA / PowerShell 免安装版
- Python 网页版，适合本机调试和扩展

它当前针对的真实流程是：

1. 打开 `5+7.xlsx` 这类源预测文件
2. 在 `Sales-CB` / `Sales-CV` 里按产品和值筛选
3. 对指定列求和，例如 `BZ`
4. 把结果写入 `Welding-CTOH` / `Welding-LV` 模板中的指定产品行和列

## 预览

![CU导入工具界面](docs/assets/cu-tool-ui.png)

## 适用场景

- 车型很多，但规则一致
- 需要人工从源 Excel 汇总后写到模板
- 不方便在 Windows 机器上安装 Python
- 需要先预览，再写入，减少误改

## 快速开始

### 方案一：直接打开 HTML

1. 打开 `excel_mapper_browser.html`
2. 先上传源文件，再上传目标模板
3. 填写规则行
4. 点击 `预览结果`
5. 确认无误后点击 `生成更新文件`

### 方案二：Windows 免安装

1. 双击 `run_windows_html.bat`
2. 或双击 `excel_mapper_html.hta`
3. 按页面提示选择源文件和目标文件

### 方案三：Python 网页版

```bash
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\streamlit.exe run app.py --server.address 127.0.0.1 --server.port 8501
```

## 规则说明

每一行规则对应一次写入，核心字段如下：

- `源Sheet`：例如 `Sales-CV`
- `源产品列 / 值`：例如 `A / CAT`
- `源部件列 / 值`：例如 `D / C7.1 TA`
- `求和列`：例如 `BZ`
- `目标Sheet`：例如 `Welding-CTOH`
- `目标客户列 / 值`：例如 `D / CATWuxi`
- `目标部件列 / 值`：例如 `B / C7.1TA`
- `写入列`：例如 `N`

匹配规则：

- 忽略大小写
- 忽略空格、换行和多余空白
- 客户值和部件值都要命中
- 求和列里的非数字会忽略并写日志

## 常见约定

- 只支持 `.xlsx`
- 先 `预览结果`，再 `生成更新文件`
- 浏览器版不会直接覆盖原文件，只会下载新文件
- 目标表找不到行时，先修正规则再执行

## 文件结构

- `excel_mapper_browser.html`：单文件 HTML 版
- `excel_mapper_html.hta`：Windows HTA 版
- `excel_mapper_windows.ps1`：PowerShell 版
- `app.py` / `excel_tool.py`：Python 网页版
- `tests/`：基础单元测试

## 开源许可

本项目使用 MIT License，见 [LICENSE](LICENSE)。

## 作者

制作人：Huo Yu’an  
Email：yuan.huo@tenneco.com
