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

仍需集成方补核的增量是：本快照之后生成的月表最终 `result.png`、`execution.json`／运行总结，最终发布报告、下载验证、CI 证据，以及扫描后继续修改的 README／发布文档。只需检查新差异和图片，不必重复业务处理或历史全扫。本报告自身在扫描后补入了统计及人工分类文字；没有引入业务值。

## 可公开内容与边界

现有诊断材料会公开本地用户名／路径、业务报表文件名和工作表名称、行格及合并数量、文件哈希、构建环境和提交标识。这些是工程诊断元数据；本轮授权范围没有要求抹去这些信息，未把它们自动判作凭据或业务原值。真实原件及派生工作簿仍应留在 ignored `testfile/`，可执行文件和交付 ZIP 留在 ignored `dist/` 并通过正式 Release 资产交付。

当前未发现需要秘密清除或历史重写的具体对象。模式匹配不能证明不存在任何未知凭据；本检查未连接远端，不覆盖远端独有引用、不可达历史对象、LFS 服务、Release 资产、Issue／PR 评论或外部附件。图片采用人工查看而非 OCR。最终公开动作应绑定通过增量核对的实际提交，不能把本快照作为后来新增文件已审查的证明。
