# 最终流式版本独立内容验收

**已提供的六份实际文件，两种真实 EXE 模式的 12 份输出，以及双大表派生负载全部通过独立内容校验。未发现本轨道的发布阻断项。** 此结论不包含尚未收到的约 100MB 月度报表，也不将派生复制样本称为真实月报。

冻结引擎 SHA256：`1d761a9f897b545c367514069e755f1aa270f1d9a777d963ffe1af41b606b5eb`。

候选 ZIP SHA256：`07b801517060896c01ea9be2aaf1420784d2e87fe3a5f2ae4ea1b31f5b0488f6`；实际 GUI 执行的 EXE SHA256：`239b1456771ae417f044dfd1ac2642f7ac36de7bc3781c6d2ac3ddc75d2d7ea4`。构建来源见 [清单](build/attempt-2/verification-manifest.json)。

| 检查 | 结果 | 证据 |
|---|---|---|
| 最终引擎 DOM / 流式差分 | 16 个合成场景 × 2 模式全部等价；ZIP 元数据、无匹配表原字节、跨表失败清理、临时文件关闭通过 | [独立代码审查](streaming-independent-review.md)、[原始结果](stream-differential-review.json) |
| 真实日明细最终源码默认输出 | 与已全量核对 2,667,926 格的首轮结果逐字节一致，原件和引擎身份匹配 | [继承闭环](streaming-final-content-inheritance.json)、[首轮全量报告](streaming-first-content-validation.json) |
| 实际 EXE 默认模式 | 6 / 6 输出通过，共 2,680,287 格，0 问题 | [汇总](exe-content/default/summary.json) |
| 实际 EXE 全部拆分 | 6 / 6 输出通过，共 2,680,309 格，0 问题 | [汇总](exe-content/all/summary.json) |
| 双大表派生负载 | 3 表共 5,335,846 格全量通过，0 问题 | [全量报告](streaming-multisheet-content-validation.json) |

双大表负载保留原 2 行小表，并复制真实日明细主表一次，因此实际为 **2 张各 190,566 行的大表 + 1 张 2 行小表，共 381,134 行**。独立 oracle 从源合并锚点推导 496 个目标合并、1,523,996 个填充格，逐行核对全部输出内容、类型、样式、行属性和顺序、合并集合、dimension 及非目标 ZIP 部件；源与结果在校验前后均未改变。校验耗时 59.95 秒，仅代表独立核对的时间。

该派生源 SHA256 为 `c55d95aa6d7f46b5521040b740207e675357c4eef72b9e91c911454b8344bc81`，输出为 `2f0ecb626f951677093a81037eaf74e56452e600c920db4597dcc7ed8e7413c4`，与 [处理测量报告](measure-streaming-multisheet.json) 完全匹配。处理耗时 129.87 秒，峰值 RSS 71.98 MiB / 私有内存 64.66 MiB；单主表最终测量为 63.46 秒、RSS 71.92 MiB / 私有内存 64.20 MiB。该样本中没有观察到跨表累计内存；这不是对任意文件大小和结构的内存上限承诺。

真实 GUI 完成界面与无匹配再次处理由 [GUI 验收报告](exe/review.md) 记录；本报告负责输出内容的独立核对，完整 12 份输出统计见 [业务内容报告](exe-content/report.md)。所有校验只读取业务文件，证据不保存单元格业务值。
