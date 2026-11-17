# Excel 语义变更跟踪工具

这是一个 Windows 桌面程序：选择旧版和新版 Excel 工作簿，按 Item Type 筛选记录，用本地多语言语义模型匹配改写前后的需求，再生成带字段对照的 Excel 报告。

程序使用 Python、Tkinter、openpyxl 和 Sentence Transformers。语义模型在本机运行，不需要付费 API 或 API Key。阈值附近或存在多个相似候选的记录不会被强行配对，可在 GUI 中确认、拒绝，或输入另一条 After 原始行号。

## 另一台 Windows 电脑上安装

### 适用环境

- Windows 10 或 Windows 11，64 位
- Python 3.11 或 3.12，64 位
- 首次安装依赖和首次下载语义模型时需要联网
- 本项目仓库为私有仓库；请使用已获仓库访问权限的 GitHub 账号下载

### 下载和启动

1. 打开 GitHub 仓库页面，点 **Code → Download ZIP**，下载并解压项目文件夹。
2. 安装 Python 3.11 或 3.12。安装界面中启用 **Add Python to PATH**，并保留 Python Launcher（`py`）。
3. 双击项目文件夹里的 **`run.bat`**。
4. 第一次启动时，脚本会创建项目专属的 `.venv` 环境并安装依赖。窗口会显示安装进度；安装完成后 GUI 会自动打开。
5. 第一次启用语义匹配时，程序会下载 `paraphrase-multilingual-MiniLM-L12-v2` 模型文件。下载时间取决于网络。模型保存在 Hugging Face 本机缓存中。
6. 以后双击 `run.bat` 即可启动。依赖和模型缓存仍在本机时，后续比较可以离线运行。

安装成功的直接标志是桌面 GUI 窗口打开，并能看到 Before/After 文件选择、工作表和列设置。如果想进一步验证项目代码，可在项目文件夹运行下面的自动化测试命令。

如果已经安装了依赖，也可在项目文件夹中运行 `python main.py`。需要手动安装依赖时，在项目文件夹运行 `python -m pip install -r requirements.txt`。

## GUI 操作顺序

1. 浏览并选择 Before 和 After 的 `.xlsx` 或 `.xlsm` 文件。
2. 分别选择两份文件中的工作表。程序会读取第一行标题，并显示列字母、编号和标题。
3. 分别选择内容列和 Item Type 列。标题为 `ObjectText` 或 `Item Type` 时程序会尝试推荐对应列；仍应检查两侧选择。
4. 输入筛选值，默认是 `Requirement`。筛选忽略大小写及首尾空格。
5. 按需调整匹配阈值（默认 0.68）。值越高，自动确认越保守。默认开启本地语义匹配。
6. 选择输出文件夹，点击 **开始分析并生成报告**。
7. 对有歧义的候选，选择候选、输入其他 After 行号、确认不匹配，或暂不处理。暂不处理的关系不会被自动配对，会保留在报告的 `Review_Candidates` 工作表中。
8. 报告默认保存为 `Change_Tracking_Result.xlsx`。GUI 显示分类计数，并提供打开结果和输出文件夹的按钮。

新旧工作表标题或顺序不一致时，GUI 会提醒报告字段默认按物理列号配对（A 对 A、B 对 B）。请确认后再继续。

## 报告内容

- `Change_Tracking`：依次排列 `changed`、`new`、`out`、`still`，包含 Before/After 原始位置及所有字段的并排值。
- `Summary`：统计分类数、筛选记录数、匹配数、人工确认数和待复核候选数。
- `Review_Candidates`：仅当存在暂未确认的候选时出现，保留旧内容、候选新内容、原始行号、匹配分数及方法。

新增记录的 Before 字段、删除记录的 After 字段以数值 `0` 表示；真实空单元格保留为空。行号变化不作为字段变化。模型分数只是候选依据，低分或歧义记录需人工判断。

## 验证

在安装好项目依赖后，从项目文件夹运行：

```text
python -m unittest discover -s tests -v
```

测试通过的情况下会显示 `OK`。测试使用临时工作簿，不会修改用户的 Excel 文件。

## 更新程序

再次下载新版本 ZIP 并解压到新文件夹，或使用 Git 更新仓库；然后双击新版本中的 `run.bat`。若 `requirements.txt` 有变化，删除新版本目录中的 `.venv` 后再次运行 `run.bat`，脚本会重新安装依赖。模型缓存可保留。

## 常见问题

- **提示找不到 Python**：安装 64 位 Python 3.11/3.12，并启用 Python Launcher 或 PATH，然后重新双击 `run.bat`。
- **依赖安装失败**：检查网络后重试。脚本会保留已创建的 `.venv`；若环境不完整，可删除 `.venv` 再运行。
- **模型下载或加载失败**：首次使用语义匹配时联网重试；若模型已缓存，检查缓存文件是否完整。也可以先关闭语义匹配，仅运行完全一致匹配。
- **Hugging Face 提示 Windows 符号链接不可用**：不是错误，模型缓存仍能正常使用，但可能占用更多磁盘空间。
- **筛选不到记录**：确认 Item Type 列选择正确，并检查筛选值是否与单元格内容相符（忽略大小写和首尾空格）。
- **列标题顺序不同提示**：确认新旧文件的物理列对应关系。报告按相同列号输出，用户选定的内容列也分别配置。
- **结果文件被占用**：关闭 Excel 中已经打开的 `Change_Tracking_Result.xlsx` 后重新生成。

## 项目结构

```text
main.py                 Tkinter 启动入口
app.py                  Windows GUI、后台任务和人工复核界面
models.py               数据结构
excel_reader.py         工作簿/工作表读取及筛选
semantic_matcher.py     本地语义向量、候选召回、一对一匹配和复核
change_detector.py      原始字段比较和分类
excel_writer.py         格式化 Excel 报告
requirements.txt        运行依赖
run.bat                 Windows 双击启动脚本
tests/                  自动化测试
```

