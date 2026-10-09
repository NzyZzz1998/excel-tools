# v1.1 发布包独立验收

2026-10-09。本轨道只读核对候选与 Git/CI 身份，并从现有 ZIP 新解压运行一次隐藏自检；没有重新打包、修改应用源码、读取业务工作簿、创建 tag/Release 或上传制品。

**结论：包与源码身份、内容、自检及指定 CI 核验通过，无已确认的发布阻断项。** 后续发布者仍需核对实际 tag 目标、上传资产名称与上传后下载哈希；本记录不提前宣称尚未创建的 Release 已经可下载。

用户随后追加月度大数据性能要求，整体发布已暂停等待该轨道结论。本报告只确认上述既有候选的包验收，不代表新增性能要求已通过；若应用实现变化，需要对新候选重新构建并绑定验证身份。

## 候选和源码身份

| 对象 | 本轮重新核对结果 |
|---|---|
| 基准提交 | `96cfbfeab30be280129aa8af9fff74e5f6be2464` |
| 本地候选 | `dist/ExcelTools-v1.1-Windows-x64.zip`，12,162,296 字节 |
| ZIP SHA256 | `92870bdc52451341c7268f6e1937a3741ecd3e1012de0cf2e4026c606ccdd4a9` |
| EXE SHA256 | `066ca4a1c7fe7fd2ac80a06ddc3969dbd178f5181983d8767d2c219078485d5c` |
| 格式 | PE x64 (`0x8664`)，Windows GUI subsystem (`2`) |
| 包成员 | 仅 `ExcelTools.exe`、`使用说明.txt`；CRC 通过，无绝对路径/越界成员 |
| 说明 | UTF-8 BOM，正文逐字节等于已验收 `packaging/使用说明.txt` 的 BOM 转换，显示 v1.1 |

引擎、GUI、4 份仓库测试及包内说明来源共 7 个文件，当前工作区字节哈希均与业务修复后的冻结构建清单一致。Git 的 LF blob 与 Windows CRLF 工作区分别记录哈希；统一 CRLF/LF 后均相等，不把换行差异当成源码漂移。

进一步只读解析 EXE 的 PyInstaller archive，抽取 `excel_unmerge_gui` 与 `excel_unmerge_fill` 的 Python 代码对象，与 Python 3.12.8 编译该 Git 提交所得代码递归比较。仅排除构建路径 `co_filename`，字节码、常量、参数、名称、行号表、异常表均相等。此证据直接连接二进制内部应用实现与提交，不仅依赖历史文件名。EXE archive 未发现业务工作簿、CSV、7z、环境凭据文件等额外数据成员。

证据：[package-audit.json](package-audit.json)、[可复核脚本](audit_package.py)。原始构建清单位于 [fixed-build](../../business-acceptance-2026-10-09/fixed-build/verification-manifest.json)。如最终 tag 包含后续文档提交，应用源码应继续与本基准相等，并由发布者记录实际 tag 目标。

## 实际可执行文件自检

本次从上述 ZIP 新解压到含中文和空格的临时目录，独立工作目录，删除 Python/Tcl/虚拟环境变量，限制 PATH，使用独立 TEMP/TMP 后通过 `Start-Process -WindowStyle Hidden` 运行 `--self-test --self-test-log`。

- 实际运行 2026-10-09 01:29:51–01:29:53 UTC；退出码 0，自己启动的进程已退出。
- 实际日志为 `ExcelTools v1.1 self-test passed.`，与包内说明及源码窗口版本一致。
- 自检执行两种拆分规则、输出内容与合并检查、原文件保护、真实 Tk 控件布局及结果状态、预停止状态和错误中文指引。它不等价于全部用户操作或全部办公客户端验收。
- 运行后候选 ZIP 哈希不变。Authenticode 状态是 `NotSigned`，未作已签名声明。

证据：[本次过程记录](fresh-portable-process.json)、[实际日志](fresh-portable-self-test.txt)、[启动脚本](run_package_self_test.ps1)。业务修复后历史自检记录也为退出 0，本次已重新验证对应候选。

## 远端 CI 实际结果

[GitHub Actions run 37866755981](https://github.com/NzyZzz1998/excel-tools/actions/runs/37866755981) 为 main 的 push 触发，`headSha` 精确等于上述基准提交，最终 conclusion 为 success。

| 步骤 | 实际记录 |
|---|---|
| 源码测试 | `Ran 43 tests in 5.345s`，`OK`，包含 7 个真实 Tk GUI 测试 |
| 依赖 | Python 3.12.10；PyInstaller 6.22.3；openpyxl 3.1.5 |
| 构建 | Windows Server 2022 x64 runner，onefile/windowed，Build complete |
| 隔离自检 | 实际步骤 success；日志展示中文/空格目录、清环境、隐藏启动和退出码检查命令 |
| 失败诊断 | 因前序成功而 skipped，符合只在失败时收集的设计 |
| 用户包准备及上传 | 均 success；artifact ID `11588158760`，12,130,589 字节 |

证据：[API 步骤状态](ci-run.json)、[完整步骤日志](ci-steps.txt)。日志第 216–218 行为测试总结果，第 229–230 行为构建版本，第 298–336 行为隔离执行命令，第 374–377 行为上传身份。成功时 CI 自检日志没有单独上传；其成功依据是步骤退出状态及脚本对非零退出抛错的检查。本地候选另有可读的实际成功日志。

CI 使用 Python 3.12.10，本地已验收包使用 3.12.8。CI artifact digest 为 `10ec285b9bef04e5940cdf74aec717baec74e7a828fbe26fa7b11005a27100b0`，与本地候选不同；本轮未下载或替换候选，也不宣称两次构建逐字节相同。CI 证明同提交能在另一 Windows 环境测试并构建成功。

## 发布文档和下载链路

当前 README 的 latest 下载资产名是 `ExcelTools-Windows-x64.zip`。计划 tag 为 `v1.1` 时，将本地候选按该资产名上传即可保持该入口；仅改外部资产名不改变 ZIP 内容或 SHA256。发布前已通过认证 API 确认既有 latest 是 `v1.0.0`，其同名资产存在，见 [原 latest 元数据](previous-latest-release.json)。新 Release 必须为正式发布并成为 latest，发布后再核验固定 tag 下载与 latest 的实际结果。

初始文档检查提出三项小改，已交由发布者处理：

1. README 第 26 行“本地候选”及发布说明的“发布前应核对”措辞，正式发布后改成已执行事实。
2. 说明升级时先正常退出旧程序，把新版完整解压到独立目录再运行；当前任务记录仅保留在窗口，不迁移未完成任务。
3. 加入 [v1.0.0 固定版本入口](https://github.com/NzyZzz1998/excel-tools/releases/tag/v1.0.0) 和回退步骤，保留原始 Excel，不覆盖用户文件；旧版不具备 v1.1 修复。

仓库访问权限保持原样；链接可用性应在当前账户授权范围内验证。此轨道不修改仓库可见性。干净 Windows 10/11、WPS、完整人工视觉检查等既有边界不因本次 CI 成功而自动完成。

## 性能对照基线核对

应月度性能轨道要求，额外只读核对了 `v1.0.0`。该 tag 指向 `470a4e1b3c5346ed8bdfd3d770b4d235eb45f173`，其 `excel_unmerge_fill.py` Git blob 为 `9d903d64caec59922fc1febf339ec0ab71dfd969`，LF 原始字节 SHA256 为 `348796863e6829bbf529f2b4629d7cca64fbab777af1a043f0af1b0fb6a40f21`。

既有 `evidence/baseline_engine.py` 的 SHA256 为 `252bbf1edfe9c24d78d4d3531b6d0be31c20f395bd79a6bdf171eb47a0690528`，仅多出 304 个 CRLF 中的 CR；统一换行后与 tag 引擎逐字节相同，可作为正确的 v1.0 性能基线。证据：[baseline-identity.json](baseline-identity.json)。本轨道没有执行性能对照。
