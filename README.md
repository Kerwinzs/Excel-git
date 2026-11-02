# ExcelGit：用 Git 管理 Excel 的小练习

这是一个本地 Python 命令行项目。你继续用 Excel 编辑原文件，用这个工具给文件保存版本、查看历史、导出旧版，以及生成单元格差异 CSV。

没有服务器、账号、数据库或网页。所有 Git 操作都在本地进行，不会自动上传文件。

想了解项目架构、完整执行流程和关键代码，请阅读 [快速上手项目指导](快速上手项目指导.md)。

## 1. 安装与运行

需要 Python 3.10+ 和 Git。此项目已在 Python 3.12.4、Git 2.39.5、openpyxl 3.1.5 下验证。

在终端进入**本 README 所在目录**，执行：

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -e .
python -m excelgit --help
```

Windows 使用 `python -m venv .venv`，PowerShell 中用 `.venv\Scripts\Activate.ps1` 激活，然后执行同样的安装命令。系统必须能在终端中执行 `git --version`。

安装后的 `excelgit` 与 `python -m excelgit` 等价。下文统一使用后者，便于理解实际运行的 Python 包。

如果当前 Python 已有 openpyxl，可以先不安装项目，直接在项目目录执行 `python3 -m excelgit --help` 和测试。

## 2. 五分钟跑完一轮

以下命令从项目目录执行。示例会创建两份练习文件，不需要你的真实 Excel。

```sh
# 生成初版和修改版：采购数量、公式和工作表有变化
python examples/make_samples.py --output demo/input

# 在独立的新目录创建版本库
python -m excelgit init demo/repository

# budget 是工作簿的稳定名称；-m 是这一次快照的说明
python -m excelgit snapshot demo/input/budget_v1.xlsx --name budget -m "初版：录入采购" --repo demo/repository
python -m excelgit snapshot demo/input/budget_v2.xlsx --name budget -m "核对数量并更新公式" --repo demo/repository

# 看这个工作簿的历史，最新版本排在最上方
python -m excelgit history budget --repo demo/repository

# 此示例恰有两个快照，HEAD~1 是初版，HEAD 是修改版
python -m excelgit diff budget HEAD~1 HEAD -o demo/reports/changes.csv --repo demo/repository

# 导出初版副本，当前文件不变
python -m excelgit export budget HEAD~1 -o demo/export/budget_original.xlsx --repo demo/repository

# 也可以跳过 Git，直接比较两个磁盘文件
python -m excelgit compare demo/input/budget_v1.xlsx demo/input/budget_v2.xlsx -o demo/reports/files.csv
```

打开 `demo/reports/changes.csv` 可以看到 6 条变化：数量和公式各改一处、删除「备注」表及其 A1、新增「审批」表及其 A1。报告含中文表头，使用带 BOM 的 UTF-8 编码，可直接用 Excel 打开。

命令重复执行时，如果示例文件、导出文件或报告已存在，工具会报错。继续练习时换一个输出路径即可。同一个原始文件的内容完全相同时，再保存快照不会重复提交。

## 3. 管理你自己的文件

先为版本库选一个**新文件夹**，原始文件放在它外面。例如：

```sh
python -m excelgit init "./my-excel-history"
python -m excelgit snapshot "/你的目录/报表.xlsx" --name "月度报表" -m "11月初版" --repo "./my-excel-history"
```

之后仍在 Excel 中编辑 `/你的目录/报表.xlsx`，保存完成后，再运行 `snapshot`，使用相同的 `--name "月度报表"` 和新的提交说明。不要直接编辑版本库中的托管副本。

一个版本库可以管理多个工作簿。`--name` 不带 `.xlsx`，允许中英文、数字、下划线、空格和短横线，长度最多 80 字符。

`history` 会显示完整提交哈希、时间和说明。日常使用时，从这份历史中复制提交哈希来导出或比较，可使用唯一的哈希前缀。`HEAD` 是整个仓库的最新提交，`HEAD~1` 是上一个提交，**不一定是某个工作簿的上一个版本**，尤其是多个工作簿共用仓库时。

初始化会有一个只保存工具配置的提交，它不会出现在工作簿的 `history` 中。在初次工作簿快照之前，该工作簿还不存在，不能导出。

要把旧版继续作为新的工作基础：先 `export` 到一个新文件，用 Excel 打开确认，再把这份文件作为下一次 `snapshot` 的来源，沿用原工作簿名称即可。工具不自动替换你的当前文件。

## 4. 原文件如何受到保护

- `init` 只接受新目录、空目录或本工具已经创建的版本库。它不会接管现有普通 Git 仓库。
- `snapshot` 读取原文件字节并复制进 `workbooks/<名称>.xlsx`；不会用 openpyxl 重新保存原件。
- 已提交且没有本地修改的托管副本会在下一次快照中更新，旧字节仍保存在 Git 历史中。
- 仓库有任何未提交修改、暂存内容或未跟踪文件时，拒绝快照，以免覆盖人工改动或夹带无关文件。请先把这些文件移出/备份，再处理仓库状态。
- 导出文件和差异报告使用独占创建模式，路径已存在就报错，没有 `--force` 覆盖选项。
- 保存快照时使用仓库锁，提交失败会恢复本次操作前的托管副本和暂存状态。请避免同时用其他 Git 工具修改这个仓库。
- 仓库的默认提交身份是 `ExcelGit Local <excelgit@localhost>`，只写入这个新仓库的本地配置。可在初始化时通过 `--user-name`、`--user-email` 指定自己的身份，不改全局 Git 配置。

示例目录结构：

```text
demo/
  input/                    ← 你的练习原件，在版本库之外
    budget_v1.xlsx
    budget_v2.xlsx
  repository/               ← 独立的 Git 版本库
    .git/                   ← 历史内容和提交信息
    .excelgit.json          ← 工具识别标记
    .gitattributes          ← 不对 .xlsx 做换行转换
    workbooks/
      budget.xlsx           ← 最近一次已提交的托管副本
  export/                   ← 导出的历史副本
  reports/                  ← CSV 差异报告
```

把导出文件和报告放在版本库之外，可以避免它们变成阻止下一次快照的未跟踪文件。

## 5. 架构与关键代码

源代码只有三层，按以下顺序阅读即可：

| 文件 | 职责 |
| --- | --- |
| `excelgit/cli.py` | `argparse` 把终端命令分派到对应功能，统一输出错误 |
| `excelgit/repository.py` | 调用 Git，保存、列出和读取原文件历史 |
| `excelgit/compare.py` | openpyxl 读取内容，比较单元格，输出 CSV |
| `excelgit/common.py` | 共用错误类型，以及拒绝覆盖的新文件写入 |
| `tests/` | 内存工作簿夹具、比较测试、真实临时 Git 仓库与命令行测试 |
| `examples/make_samples.py` | 生成两份可在 Excel 中打开的练习文件 |

### A. 为什么 Git 和 openpyxl 要分工？

`.xlsx` 内部是一个 ZIP 包。Git 可以精确保留它的每个版本，但通常不能直接展示有用的文本行差异。因此我们让 Git 保存原始字节，让 openpyxl 解释工作表和单元格。

保存时，核心步骤相当于：

```python
# 已完成校验、复制及安全检查后
self.git("add", "--", relative)
self.git("commit", "--quiet", "-m", message, "--", relative)
```

一次 `commit` 就是一份可以找回的快照，提交说明是人读的版本备注，哈希是精确定位版本的标识。这里不引入标签、分支或远程仓库，先把这条路径学清楚。

`Repository.git()` 使用 `subprocess.run()` 的参数列表，传入 `capture_output=True`，读取 Git 返回内容；返回码非零时转成可读错误。没有拼接 shell 命令，即使文件名含空格，也作为一个参数传递。

### B. 为什么导出不用 checkout？

`version_bytes()` 先把用户输入的版本解析成提交哈希，再执行类似下面的操作：

```python
self.git("show", f"{commit}:{path}")
```

这会取出某个提交中某个文件的字节，不切换分支，也不替换工作目录文件。`export()` 再把这些字节写入一个新路径。

`new_output()` 使用 `"xb"` 或 `"x"` 模式。与普通写入模式 `"wb"` / `"w"` 不同，目标存在时会失败，这就是避免覆盖的关键。

### C. 怎么比较单元格？

读取时保留公式：

```python
load_workbook(BytesIO(data), read_only=True, data_only=False)
```

如果设置 `data_only=True`，公式单元格通常得到的是上次保存的缓存结果，可能为空，也无法可靠识别公式文本变化。本工具不会计算 Excel 公式。

`_contents()` 把每个普通工作表整理为：

```python
{"采购": {"B2": CellValue("n", 3), "D2": CellValue("f", "=B2*C2")}}
```

`compare_workbooks()` 比较两侧工作表名的并集，再比较单元格坐标的并集。旧侧缺少是新增，新侧缺少是删除，两侧都有但类型或值不同是修改。比如数值 `1`、文本 `"1"`、布尔值 `TRUE` 会区分。

工作表名称相同才会对应，单元格按 A1 坐标对应。插入一行后可能出现多条变化；它没有业务主键，也不会猜测整行移动。

### D. 报告怎么读？

| 变化代码 | 含义 |
| --- | --- |
| `sheet_added` / `sheet_removed` | 工作表新增 / 删除，单元格列为空 |
| `cell_added` / `cell_removed` | 单元格内容新增 / 清空 |
| `cell_changed` | 同一坐标的类型或值改变 |

新增/删除工作表还会逐格列出其非空内容。空表的增删也会有记录。常见类型是 `n` 数值、`s` 文本、`b` 布尔、`f` 公式、`d` 日期时间、`e` Excel 错误。

文本值用 JSON 字符串表示，因此报告中的公式会显示为 `"=B2*C2"`，文本数字会显示为 `"1"`，区别于数值 `1`。这种可见的双引号是报告格式的一部分；单元格中的换行会显示成 `\n`。日期时间以 ISO 格式显示。可能被 Excel 当作公式的字段会加前导单引号，以纯文本方式展示。

## 6. 运行测试

```sh
python -m unittest discover -s tests -t . -v
```

使用 Python 自带的 `unittest`，不需要额外的测试框架。夹具通过 openpyxl 在内存中创建；Git 测试只操作临时目录，不使用你的仓库或真实文件。

覆盖内容包括公式文本、空值与零、类型区别、日期、工作表增删、格式忽略、CSV 中文与特殊字符、完整快照历史、字节一致的历史导出，以及覆盖保护、脏仓库保护和提交失败回滚。

## 7. 第一版的边界

- 只处理 `.xlsx`，不处理 `.xls`、`.xlsm`、加密工作簿或损坏文件。
- 差异只覆盖普通工作表名、单元格值与普通公式文本。不比较字体、颜色、数字格式本身、批注、合并区域、图表、图片、行列宽度、工作表顺序和可见性，也不做自动合并。样式中的日期格式可能影响 openpyxl 对值类型的识别。
- 暂不比较数组公式和数据表公式，遇到时明确报错。这类文件仍可按原始字节保存快照和导出。
- 工作表重命名会显示为删除旧表并新增新表。没有变化的比较仍生成只有表头的 CSV。
- 比较面向小文件，单个工作簿声明的矩形扫描区域累计上限是 100 万格，过大的空白格式区域也可能触发限制。快照本身没有此比较上限，但仍需能由 openpyxl 打开。
- 保存的是磁盘上已写入的内容；先在 Excel 中完成保存，再运行快照。请勿同时编辑正在被读取的文件。
- 版本库存储每次提交的工作簿，长期保存大量大文件会增长磁盘占用。Git 历史与原件都在同一台机器时，不能代替设备外备份。
- 工具锁只协调本工具的快照。意外关机后如果有遗留锁，确认没有快照进程运行，再删除版本库的 `.git/excelgit.lock`；若 Git 报 `.git/index.lock`，同样先确认没有 Git 进程，再处理。发现未提交文件时先备份，避免用强制重置来清理。

理解这版以后，可以先自行加一个 `list` 命令列出已管理工作簿，练习扩展 `argparse` 与 Git 查询。先保持命令行项目的小范围。

## 参考资料

- [Git 官方文档](https://git-scm.com/docs/git)：命令列表及参数说明。
- [git show](https://git-scm.com/docs/git-show)：读取历史对象内容。
- [openpyxl 官方教程](https://openpyxl.readthedocs.io/en/stable/tutorial.html)：加载工作簿及 `data_only` 等参数。
