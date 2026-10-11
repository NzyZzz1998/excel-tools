# OOXML 文本转义与 openpyxl 校验差异

2026-10-11：已确认首个摘要差异来自 openpyxl 3.1.5 对 `inlineStr` 的解释。本项没有修改应用代码，也没有把原始失败报告改成通过。实际全表是否保真，仍以修正后的独立全量比较结果为准。

## 实际单格证据

首差坐标为 `F246`，源与输出均为文本、`General` 数字格式。源 XML 中一个 `_x003D_` 在输出中表示为等号。openpyxl 将源中的七个字符作为字面内容，因而报告字符串长度为 32；输出长度为 26。

独立 Excel 16.0 对实际源和第一个拆分结果的 `F246` 只读读取确认：两边 `Value2` 完全相等、长度均为 26、都没有公式、数字格式相同。该证据仅含地址、类型和统计，见 [excel-escape-difference.json](../excel-escape-difference.json)；这里不保存业务值或原工作簿。

## 六个纯合成案例

[create_fixture.py](create_fixture.py) 直接构造微型 OOXML 源文件，然后调用冻结的应用 reader/writer 生成候选输出。预期值逐项手写，没有使用产品解码函数推导预期。

| 案例 | 源文本表示 | Excel 中预期值 |
| --- | --- | --- |
| 普通等号 | `left_x003D_right` | `left=right` |
| 以等号开头的文本 | `_x003D_1+1` | 字面文本 `=1+1`，无公式 |
| 受保护的字面转义串 | `_x005F_x003D_` | 字面文本 `_x003D_` |
| Xstring 回车 | `left_x000D_right` | 中间为 U+000D |
| XML 字符实体回车 | `left&#13;right` | 中间为 U+000D |
| UTF-16 代理对 | `_xD83D__xDE00_` | U+1F600 |

[validate_excel.ps1](validate_excel.ps1) 创建独立 Excel 实例，并用窗口句柄对应的进程编号确认其不属于运行前已有进程；编号只在内存中用于所有权确认，不写入报告。禁用宏、外链更新和事件，只读打开两本合成工作簿，关闭时不保存。

公开版脚本已在新临时目录重新执行：源与输出共 **12 次 Value2 比较全部符合手写预期**，`HasFormula` 全为 false，两本文件的 SHA 前后未变，关闭过程无错误。见 [excel-result.json](excel-result.json)。报告不含进程编号、私有路径或业务值；工作簿保留在 ignored `testfile/` 中。

同一源和输出经 openpyxl 3.1.5 读取可复现差异，见 [openpyxl-result.json](openpyxl-result.json)。其 `WorkSheetParser.parse_cell` 的 `inlineStr` 分支通过 `Text.from_tree(child).content` 读取文本，没有执行 ST_Xstring 转义解码。XML 自身的数字字符实体会由 XML 解析器处理，因此回车实体这一项与其他编码项表现不同。

## 规范依据及校验边界

Microsoft 的 [Spreadsheet Text 文档](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.text?view=openxml-3.0.1) 说明 `t` 的内容类型是 ST_Xstring，且它可位于 `is`、`r`、`si` 等节点。Microsoft 的 [ST_Xstring 实现说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/d34ae755-c53f-4a44-a363-c6dd3ee018a4) 说明标准没有限定哪些字符可转义，并特别说明保护 `_xHHHH_` 字面串时对前导下划线的处理。上述页面于 2026-10-11 查阅。

该月表已盘点为 `inlineStr` 字符串；独立校验器可对此限定场景实施单遍、非递归的 ST_Xstring 语义解码，同时规范源和输出。`_x005F_x003D_` 必须只解码成字面 `_x003D_`，不能继续变成等号。

不能把这一补偿无条件套到 openpyxl 已读取的共享字符串（SST）：它的 `read_string_table` 已先移除保护下划线的 `x005F_` 标记，再盲目二次解码可能改变本来应当保留的字面串。需要处理 SST 时，应基于原始 XML 状态另行设计独立校验。

## 复验

使用带 openpyxl 3.1.5 的测试 Python 环境，在仓库根目录执行；工作目录须是尚不存在的新临时目录：

```powershell
python docs/releases/v1.2.0/acceptance/xstring-probe/create_fixture.py --workdir testfile/xstring-proof-new
powershell.exe -NoProfile -ExecutionPolicy Bypass -File docs/releases/v1.2.0/acceptance/xstring-probe/validate_excel.ps1 -WorkDir testfile/xstring-proof-new
```

第二步需要本机 Microsoft Excel。脚本不会打开用户的业务文件。应用与证据绑定哈希见 [verification-manifest.json](verification-manifest.json)。运行时代码与 `measure_tables.py` 均未由本项修改。
