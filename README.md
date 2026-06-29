# Excel产品数据自动提取与写入工具

这是一个本机网页工具，用于从多个源 `.xlsx` 产品文件中按客户、零件号或项目号识别车型/产品，并把提取值或求和值写入目标 `.xlsx` 表格。既支持写入固定单元格，也支持按目标表里的车型列自动找到对应行后写入指定列，适合多个车型重复处理。

## Windows电脑使用方式

### 方式一：纯HTML版（不需要Python，推荐）

1. 把本文件夹复制到 Windows 电脑。
2. 双击 `excel_mapper_browser.html`。
3. 先上传“要改的目标Excel”，再上传“参考Excel”。
4. 配置产品规则和映射规则。
5. 点击“预览匹配”，确认无误后点击“执行吸入并下载Excel”。

这个版本是真正的 `.html` 文件，不需要 Python，也不需要安装第三方库。浏览器安全限制下，它不会覆盖原文件，而是下载一个新的 `xxx_已更新.xlsx`。

#### 多车型批量模式

在“映射规则”里：

- `产品` 填具体车型名称：只处理这一辆车型。
- `产品` 填 `*` 或 `ALL`：同一条规则自动套用到“产品规则”里的全部车型。
- `目标模式` 选“固定单元格”：写入 `目标单元格`，例如 `B2`。
- `目标模式` 选“按车型找行”：程序会先在 `目标匹配列` 里找车型/客户/零件号/项目号，再把结果写到 `目标写入列` 的同一行。
- `目标匹配列` 和 `目标写入列` 可以填写表头名称、Excel列字母或列号，例如 `车型`、`B`、`2`。
- `表头行` 用于源文件找列；`目标表头行` 用于目标文件找列。

规则会保存到当前浏览器的本地存储里。换电脑或换浏览器后，需要重新配置一次。

### 方式二：HTA免安装版（可直接覆盖目标文件）

前提：Windows电脑已安装 Microsoft Excel。

1. 把本文件夹复制到 Windows 电脑。
2. 双击 `run_windows_html.bat`。
3. 在 HTML 界面窗口里选择源Excel、目标Excel，配置产品规则和映射规则。
4. 点击“预览匹配”，确认无误后点击“执行写入”。

这个版本使用 Windows HTML 应用调用本机 Excel，不需要安装 Python。执行写入前会自动备份目标文件。

如果 `run_windows_html.bat` 没反应，也可以直接双击 `excel_mapper_html.hta`。

### 方式二-B：PowerShell免安装版（备用）

如果 HTML 应用被公司策略拦截，可以尝试双击 `run_windows_no_python.bat`。
它同样不需要 Python，但需要允许运行 PowerShell。

### 方式三：网页版本（需要Python）

1. 安装 Python 3.10 或更新版本。
2. 安装时勾选 `Add Python to PATH`。
3. 把本文件夹复制到 Windows 电脑。
4. 双击 `run_windows.bat`。
5. 浏览器打开 `http://127.0.0.1:8501`。

首次启动会自动创建 `.venv` 本地运行环境并安装依赖，可能需要几分钟。

### 方式四：命令行启动网页版本

在本文件夹打开命令提示符或 PowerShell：

```bash
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\streamlit.exe run app.py --server.address 127.0.0.1 --server.port 8501
```

## macOS/Linux 启动

```bash
python3 -m pip install -r requirements.txt
streamlit run app.py
```

浏览器打开后：

1. 上传多个源 Excel 文件。
2. 输入目标 Excel 路径，或上传目标 Excel 文件。
   - Windows路径示例：`C:\Users\你的名字\Desktop\目标表.xlsx`
   - 如果目标文件在 OneDrive 或共享盘，也可以填写对应的本机同步路径。
3. 在“产品规则”里配置产品名称、客户、零件号、项目号。
4. 在“映射规则”里配置源 sheet、识别列、取数列、目标 sheet 和目标单元格。
5. 先预览，再执行写入。

## 覆盖保存说明

如果填写了“目标文件本机路径”，程序会覆盖该文件，并先生成备份：

`原文件名.backup_YYYYMMDD_HHMMSS.xlsx`

如果只上传目标文件，程序无法知道浏览器中文件的原始路径，会提供处理后的文件下载。

## Windows注意事项

- 覆盖目标文件前，请先关闭 Excel 中打开的目标文件，否则 Windows 可能会锁定文件导致写入失败。
- 纯HTML版不会覆盖原文件，只会下载修改后的新文件，因此不需要关闭原文件。
- 免安装版需要电脑上装有 Microsoft Excel，因为它通过 Excel 自身读写工作簿。
- 如果公司策略禁止 HTA/ActiveX，`run_windows_html.bat` 可能无法启动，需要 IT 放行 HTA，或改用 PowerShell 免安装版。
- 如果公司策略禁止 PowerShell 脚本，`run_windows_no_python.bat` 可能无法启动。
- 推荐使用 `.xlsx`，不支持 `.xls` 和 `.xlsm`。
- 源文件可以多选上传；目标文件如果要直接覆盖，必须填写本机路径。
- 网页版本规则会保存到程序目录下的 `rules_config.json`。
- HTML免安装版规则会保存到 `rules_config_html.json`。
- PowerShell免安装版规则会保存到 `rules_config_windows.json`。
