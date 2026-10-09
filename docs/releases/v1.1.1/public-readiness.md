# v1.1.1 公开适宜性审查

结论：已检查的工作区文件及本地所有引用可达历史中，**未识别出必须在公开前排除的真实业务原件、业务单元格明细或凭据**。本报告不执行仓库公开、上传或历史重写；最终提交的新文件和本轮检查之后的改动仍须由集成方核对增量。

## 历史与已跟踪文件

快照 HEAD 为 `4e013c4ded2af26c6546fa16a1860fff282e5842`；本地仓库不是 shallow clone。扫描覆盖 8 个可达提交、310 个去重历史 blob、286 个历史路径，以及 284 个当前已跟踪文件，共 2,481,915 字节去重历史内容。包括历史中已删除的文件和提交消息，不只检查 HEAD。

证据：[public-history-scan.json](acceptance/public-history-scan.json)，可复查脚本：[audit_public_history.py](acceptance/audit_public_history.py)。扫描没有找到业务工作簿、CSV、压缩包、数据库、私钥类文件，历史中没有 `testfile/` 或 `dist/` 文件；凭据模式候选数为 0。

对结构扫描提示的三类单元格值记录，已人工追溯来源：

| 文件 | 人工归类 |
|---|---|
| `v1.1/release-acceptance-2026-10-09/engine/excel-edge-validation.json` | 报告显式标记 synthetic_only；生成器只构造行列样式、空白字符及坐标语法测试包，值为代码中的合成常量 |
| `docs/reviews/2026-10-08-evidence/engine-fidelity.json` | 合成单元格图片及省略坐标兼容探针，不来自业务工作簿 |
| `docs/reviews/2026-10-08-evidence/probes.json` | 合成 CR/LF 字符保真和性能探针，不含业务明细 |

`docs/reviews/2026-10-08-evidence/input-sheet.xml` 与 `output-sheet.xml` 是小型合成单元格图片回归夹具，不是业务导出的 worksheet XML。

22 个历史 PNG 路径对应 18 个去重图像 blob；全部逐张查看，内容仅为工具界面、文件路径、报表文件名／工作表名称、统计和日志，没有 Excel 网格中的业务单元格内容或凭据。每个历史图像版本均可由当前文件逐字节对应。六个 JSONL 文件也检查了字段结构，仅包含自己窗口的输入动作、进程、计时和内存统计。

## 本轮新增及修改

本轮 dirty／untracked 文件与拟保留的文本构建日志另作增量扫描，共检查 355 个当前文件，凭据候选及业务原件／压缩包路径候选均为 0。证据见 [public-current-scan.json](acceptance/public-current-scan.json)，脚本见 [audit_public_increment.py](acceptance/audit_public_increment.py)。该快照包含每个检查文件的 SHA256，便于最终暂存内容对照；它不是对最终暂存清单的冻结承诺。

新增的五张 PNG 已逐张查看：两张合成 GUI 布局图，以及月表运行的初始、已选择、处理中截图；均只有应用界面和诊断元数据。`assets/readme/merge-example.svg` 明示为虚构示例。新增 post-release 样式探针中的 `Value2`／显示文本来自 `prepare_style_fixtures.py` 构造的日期序号和样式夹具，记录也标记 synthetic only；恢复选择探针使用合成路径及临时工作簿。未把这些合成预期值误作真实业务数据。

该快照之后的月表最终截图、运行记录和本地验收总报告，已按下面的提交补审完成。未来产生的下载验证、CI 证据、正式发布记录，以及本次补审之后继续修改的文件，仍需由集成方核对新差异；不必重复业务处理或历史全扫。

## 提交 1a928e7 增量补审

本轮只读补审于 2026-10-10（北京时间）收口，对象是已提交的 `1a928e784cadf742b2f3eb4443ee7eaabc2ad95d`。该提交包含 84 个变更路径，完整 Git 树有 357 个文件；逐个读取这 84 个提交 blob 及提交消息重新执行凭据／敏感文件路径检查，没有新增凭据候选、业务工作簿／压缩包路径或 JSON 原值候选。与先前 355 文件快照比较，新增四个路径为 `monthly/execution.json`、`monthly/result.png`、扫描结果 JSON 本身及 `acceptance/report.md`；有变化的 README、月表动作／内存轨迹、进度和本报告也已补核。

- [月表完成截图](acceptance/monthly/result.png) 已逐张查看：窗口为 v1.1.1，默认规则，成功 1/1、失败 0、待处理 0；日志仅有路径、工作表名及拆分／填充数量，没有工作簿网格或业务单元格值。处理中截图也再次检查，显示读取阶段、145,051 行和已用时 00:31，没有新增敏感内容。
- [execution.json](acceptance/monthly/execution.json)、[identity.json](acceptance/monthly/identity.json) 和 [progress.json](acceptance/monthly/progress.json) 只包含文件身份、哈希、应用窗口／进程、计数和资源观测。`execution.json` 中自动生成的 `ui_conclusion` 仍为待人工查看，人工补审结论记录于本报告；`progress.json` 是完成前心跳，不应解读为最终状态。最终状态以执行记录和完成截图为准。
- 动作日志共 16 条，仅含目标／命中窗口及前台恢复字段；内存轨迹 410 条，仅含时间、PID、CPU 与内存数值。未增加业务明细记录。
- [验收总报告](acceptance/report.md) 的 403.078 秒、工作集 295.86 MiB、专用内存 231.71 MiB、结果字节数及 SHA 与执行记录相符；明确内容／Excel 校验来自与 v1.1 结果逐字节一致后的证据继承，没有冒称本轮再次逐格扫描。所有相对链接均能解析到本地文件。报告未把远端 CI、Release 或公开操作写成已完成。

本轮人工检查的关键工作区文件 SHA256（Git 文本行尾可能正规化）为：

| 文件 | SHA256 |
|---|---|
| `monthly/result.png` | `fb96f0a64ca2ee9a314e4f4447928176caf40dee768bfc2aca7f09c99fa729a3` |
| `monthly/processing.png` | `74a9a20c333b8818413afd5bd34f622880702480687ca326241aa58c009ffc9b` |
| `monthly/execution.json` | `004b5482c93128abfa4dc81aeb56c338ad00fd45c83ba43b0c3fec70a747e8c4` |
| `acceptance/report.md` | `5d9e37676dadcc6d70a7c40a5724f0194f4f653ef603acdd28ca9a434ba18476` |

补审结论：该提交没有识别出必须排除的公开内容。这里只修改公开审查报告，没有修改 README、应用、测试或原验收证据，也没有运行工作簿处理或触碰用户程序。CI 修复及其后续提交不属于此 Git SHA 的检查结论。

## 可公开内容与边界

现有诊断材料会公开本地用户名／路径、业务报表文件名和工作表名称、行格及合并数量、文件哈希、构建环境和提交标识。这些是工程诊断元数据；本轮授权范围没有要求抹去这些信息，未把它们自动判作凭据或业务原值。真实原件及派生工作簿仍应留在 ignored `testfile/`，可执行文件和交付 ZIP 留在 ignored `dist/` 并通过正式 Release 资产交付。

当前未发现需要秘密清除或历史重写的具体对象。模式匹配不能证明不存在任何未知凭据；本检查未连接远端，不覆盖远端独有引用、不可达历史对象、LFS 服务、Release 资产、Issue／PR 评论或外部附件。图片采用人工查看而非 OCR。最终公开动作应绑定通过增量核对的实际提交，不能把本快照作为后来新增文件已审查的证明。

## 集成方最终增量与实际公开

集成方补核了后续 `3e26aa9` / `96b6a6e` 的测试修正、Tk弱引用探针、合成测试日志和下载验证脚本。新内容只有测试代码及生命周期/路径诊断，没有业务原件、原值或凭据。最终CI元数据只含公开run/job URL、步骤和状态；正式下载验证只含包身份/大小、隔离选项和成功文本，进程输出为空。发布记录、README最终版本提示及其校对报告也已检查。完整CI原始作业日志留在ignored目录，没有随文档上传。

另读取远端引用，均在已审查提交链；GitHub `issues?state=all` 返回空数组（包含PR入口），仓库wiki和discussions均关闭。这些查询补充上述早期本地审查的远端边界，不外推外部附件。最终README SHA见[正式补验](readme-validation.md)。

发布和下载验证通过后，按用户授权将仓库设为public。[匿名API证据](acceptance/release-publication.json)确认仓库及Latest读取均为200，正式附件digest与本地及下载后实测一致。公开动作对应应用提交/tag `96b6a6e` / `v1.1.1`，随后文档收口仅回写这些已发生的结果；真实工作簿仍留在ignored目录。
