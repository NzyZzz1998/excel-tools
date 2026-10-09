# v1.1 实际业务验收 · 2026-10-09

入口：`/acceptance`；沿用 mypm 的 Vibe Coding 全流程。用户明确提供 `E:/codex/excel-tools/testfile/期间缺货Top20` 下文件用于验收。

**结论：本批 5 个工作簿、10 张表在默认和全部拆分两种模式下验收通过。** 验收发现并修复了工作表范围声明不完整的问题，已重新构建 v1.1，并对新 EXE 的 10 份实际业务输出完成内容、默认流式读取及本机 Excel 对照。原文件没有改动。

## 当前候选身份

| 对象 | 当前身份 |
|---|---|
| ZIP | `E:/codex/excel-tools/dist/ExcelTools-v1.1-Windows-x64.zip`，12,162,296 字节 |
| ZIP SHA256 | `92870bdc52451341c7268f6e1937a3741ecd3e1012de0cf2e4026c606ccdd4a9` |
| EXE SHA256 | `066ca4a1c7fe7fd2ac80a06ddc3969dbd178f5181983d8767d2c219078485d5c` |
| 引擎 SHA256 | `95b5455b5d10aab9436a8842bbc27f6814485c037247d8320e3c9b1cf81e6e30` |
| GUI SHA256 | `0d59d53e067da326984702ece5939347448a829012b564363eb9469c42b25f87` |
| 源码身份 | `main` / `470a4e1b3c5346ed8bdfd3d770b4d235eb45f173` 加未提交改动；实际快照见哈希清单，旧 HEAD 不能单独代表此包 |
| 构建环境 | 本机 Windows 11；Python 3.12.8、PyInstaller 6.22.3、openpyxl 3.1.5 |

证据：[构建清单](fixed-build/verification-manifest.json)、[源码快照](fixed-build/source-snapshot-hashes.json)、[包内字节与 CRC 核验](fixed-build/package-content-check.json)。旧包完整归档在 `dist/archive/ExcelTools-v1.1-pre-business-fix-4a0ca98b.zip`，旧 SHA256 为 `4a0ca98b0f95d172c35a12219e2582ee95fab39dec771e74a11ea75c2c97927c`；旧候选记录只作历史证据。

## 实际结果

以下数量包含每本的第二张表；桌面版按既有合同处理所有工作表。“填充”是新增有值单元格数量，不是总行数。

| 文件 | 默认：拆分区域 / 填充格 / 剩余合并 | 全部拆分：拆分区域 / 填充格 / 剩余合并 |
|---|---:|---:|
| 供应商缺货Top20 (1).xlsx | 4 / 24 / 1 | 5 / 25 / 0 |
| 库存满足率 (1).xlsx | 9 / 74 / 1 | 10 / 77 / 0 |
| 每日缺货Top50 (1).xlsx | 45 / 1555 / 1 | 46 / 1562 / 0 |
| 期间缺货Top20.xlsx | 8 / 160 / 1 | 9 / 164 / 0 |
| 缺货数据统计.xlsx | 13 / 141 / 1 | 14 / 144 / 0 |
| **合计** | **79 / 1954 / 5** | **84 / 1972 / 0** |

默认模式保留 5 处横向“合计”合并；全部拆分多填充的 18 格均为该文字。所有合并锚点均为文本，没有新增复制的数字。原数字、编号、文本、普通空白、样式语义及非目标 ZIP 部件保持一致；例如供应商表第一张表 `B6` 的普通空白未被误填。

原始静态汇总值不能简单等同于明细求和。本工具保留源汇总值，不重新定义业务口径；此次不将源报表本来的汇总差异判为处理错误。

## 发现与修复

输入 10 张表的 `dimension` 均声明为 `A1`，但实际数据明显超出。初始候选只按被拆合并扩展范围，造成部分普通数据仍落在声明范围外。例如每日缺货表默认结果声明为 `A1:D401`，实际物理格范围达 `A1:AA402`。XML 内的数据没有丢失，Excel 仍能读到；依赖声明范围的流式读者却可能漏读。

本次修复在已处理表上，将原范围扩展到全部实体格和保留合并；不缩小较大原范围，不改没有匹配合并的表，不引入依赖。新增 3 项纯合成回归，包含远端数值、空白横向合并、原范围保持及无匹配表原样保留。

[旧候选缺陷取证](dimension-finding.json) 记录坐标、计数及哈希，不含业务单元格值。初始 `engine-content-validation.json` 的通过只代表内容核对，不包含此次新补的范围和默认流式读取检查。

## 验证层次与证据

| 验证 | 实际结果 | 证据 |
|---|---|---|
| 独立输入盘点 | 5 本 / 10 表；79 纵向、5 横向合并；前后原文件哈希一致 | [输入清单](input-inventory.json) |
| 修复后源码业务处理 | 两种模式 10/10；期望由源合并独立推导 | [源码轨道](engine-content-validation-fixed.json) |
| 新 EXE 实际处理 | 通过真实文件选择框、拆分选项和开始按钮生成两批 10 个输出；两种模式均显示成功 5、失败 0、待处理 0，没有用源码调用代替 EXE | [GUI 证据索引](gui/gui-result.json)、[默认完成截图](gui/default-proof-result-final.png)、[全部拆分完成截图](gui/all-result-final.png) |
| 结果再次处理 | 两种模式各自选择对应 5 份已处理结果，均显示成功 0、无需处理 5、失败 0；前后文件集合及哈希一致，没有重复输出 | [默认无需处理](gui/default-no-match-result-final.png)、[全部无需处理](gui/all-no-match-result-final.png) |
| 新 EXE 输出内容 | 10/10；共核对 24,740 个物理格的 XML 语义、预期填充及保留合并；普通读取与默认流式读取一致；非目标部件字节不变 | [新 EXE 内容证据](exe-content-validation-fixed.json)、[验证脚本](validate_business.py) |
| 实际 Excel 客户端 | Microsoft Excel `16.0.17932.21000`；5 个原副本和 10 个新 EXE 输出以普通模式只读打开；两种规则的 `Value2` 对照共 24,764 个格位，包括空白，10/10 通过 | [Excel 证据](excel-client-exe-fixed.json)、[COM 脚本](validate_excel.ps1) |
| 文件保护与进程结束 | Excel 前后 25 个文件哈希一致；自建 Excel 进程已退出，没有保存文件 | 同上 |
| 完整回归与构建 | 43/43 测试，9/9 独立验收；源码与独立 EXE 自检通过 | [测试](fixed-build/unit-tests.txt)、[独立验收](fixed-build/independent-acceptance.txt)、[EXE 自检](fixed-build/portable-self-test.txt) |

源码轨道先于新包构建，其 `candidate_identity` 中记录的 ZIP 是当时仍在磁盘上的旧包；源码轨道没有执行那个 ZIP。新包验收由新 EXE 轨道及其 `92870bdc…` 身份证明，两者不可混用。

早期离屏截图存在未重绘的进度缓存，只作为过程材料；正式 GUI 结论使用 `final` 截图。默认最终截图来自独立副本补跑，5 个输出与已验收默认结果字节完全相同。自建 EXE 已正常关闭，原始文件和两批已验收输出的最终哈希均未改变。

Excel 使用独立不可见实例，`UpdateLinks=0`、禁用宏、关闭事件、只读、不保存，`CorruptLoad=0` 普通打开。此模式不请求恢复，参数语义见 [Microsoft Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)。由于 `DisplayAlerts=false`，此证据不等于人工观察“没有修复弹窗”，也不证明所有视觉布局。前两次 COM 尝试因验收脚本传递缺省参数失败，已改为明确参数；失败记录保留，新 EXE 完整成功记录才是本次客户端结论。

## 交付与边界

已验证结果单独整理至 `E:/codex/excel-tools/testfile/验收_v1.1_2026-10-09/交付结果/`，分为“默认保留横向合并”和“全部拆分”，每组 5 份。复制结果与实际 EXE 输出逐一校验哈希；详见 [最终交付一致性](delivery-integrity.json)。日常保留原报表横向汇总结构可直接使用默认组。

本批没有公式、隐藏表、宏、图片、图表、外链等样本，不能据此声称它们都完成业务客户端验收。输入缺少默认样式的读取器警告源自原文件，未擅自重建样式。WPS、干净 Windows 10/11、Excel 人工视觉检查，以及实际业务中的暂停/失败重试体验仍保留补证边界；同候选合成回归对恢复状态的验证继续有效。

所有业务文件仅在本机处理。本批验收结束时尚未提交、推送、打 tag 或上传 Release；后续用户授权的 Git 同步范围见 [进度](../progress_v1.1.md)。业务原件、压缩包及结果副本保留本地，当前便携包也不纳入 Git 提交。

Memory（auto）：已更新并回读 `Excel Tools project state`，仅保留本轮结论、候选身份和证据导航；正式事实仍以仓库及制品实物为准。
