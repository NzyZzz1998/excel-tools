# 新候选 EXE 原生 GUI 验收

验收日期：2026-10-09。结论：本轮四组实际 EXE GUI 流程全部通过，无新增阻断问题。

## 包身份与输入边界

- ZIP：`ExcelTools-v1.1-Windows-x64.zip`，SHA256 `07B801517060896C01EA9BE2AAF1420784D2E87FE3A5F2AE4EA1B31F5B0488F6`。
- 实际启动 EXE：`testfile/大文件验收_v1.1_2026-10-09/exe-check/app/ExcelTools.exe`，SHA256 `239B1456771AE417F044DFD1AC2642F7AC36DE7BC3781C6D2AC3DDC75D2D7EA4`。
- 每次启动前核对 ZIP 与 EXE 双哈希；四次均读取实际窗口标题并在 fresh 截图中看到 `Excel 数据处理工具 v1.1`。
- 固定六份原始工作簿，包含 190,566 行真实日明细；身份及 SHA 见 [identity-and-inputs.json](identity-and-inputs.json)。没有加入后续其他规模样本。
- 原件、两组输入副本与业务结果全部留在 `testfile`。本证据目录只含控制脚本、窗口截图、文件路径、哈希、计数及动作记录，没有工作簿或业务单元格原值。

## 实际结果

| 流程 | 原生界面人工核对 | 文件核验 | 观测秒数 |
| --- | --- | --- | ---: |
| 默认模式 | 横向选项未勾；已处理 6/6、成功 6、无需处理 0、失败 0、待处理 0 | 六份新结果；原件与副本 SHA 不变 | 68.609 |
| 全合并模式 | 横向选项已勾；已处理 6/6、成功 6、无需处理 0、失败 0、待处理 0 | 六份新结果；原件与副本 SHA 不变 | 68.172 |
| 默认结果重跑 | 横向选项未勾；已处理 6/6、成功 0、无需处理 6、失败 0、待处理 0 | 只选择六份默认结果；目录仍为 12 个文件，全部 SHA 不变 | 34.828 |
| 全合并结果重跑 | 横向选项已勾；已处理 6/6、成功 0、无需处理 6、失败 0、待处理 0 | 只选择六份全合并结果；目录仍为 12 个文件，全部 SHA 不变 | 34.594 |

两组无需处理截图中的可见日志均显示“拆分 0 个合并区域，填充 0 个单元格”及“未生成新文件”。没有把原件与结果混合选择，也没有仅凭文件数量推断 GUI 已结束。

四次均验证“选择文件”按钮恢复可用、打开对话框并立即取消，证明 GUI 已消费结束事件，然后强制重绘后截图。各窗口最后通过自身 `WM_CLOSE` 正常退出，launcher 退出码均为 0、残留自建 PID 均为 0，无强制终止。

计时包含结果数量稳定等待、空闲确认对话框及控制脚本开销；这是本机端到端观测，不能当作纯引擎基准或跨机器性能保证。两种输出的独立全量内容 oracle 由质量审查任务另行记录，本报告不代替其内容校验。

## 证据

- 默认：[选择截图](default/selected.png)、[结果截图](default/result.png)、[执行记录](default/execution.json)。
- 全合并：[选择截图](all/selected.png)、[结果截图](all/result.png)、[执行记录](all/execution.json)。
- 默认结果重跑：[结果截图](default-no-match/result.png)、[执行记录](default-no-match/execution.json)。
- 全合并结果重跑：[结果截图](all-no-match/result.png)、[执行记录](all-no-match/execution.json)。
- 四组目录下的 `actual-input-actions.jsonl` 记录原生鼠标操作；每次点击前必须命中指定自建 PID 的指定控件，完成后恢复鼠标位置及原前台窗口。启动默认隐藏、主窗离屏，截图仅短暂使用近透明自建窗口强制重绘。
- `native_gui_controller.py` SHA256：`43A8205FCF257BD97063404F167BC587FEE3A6A8C939FB23ABB058DE21B5073B`。
- `run_real_exe_batch.py` SHA256：`EEF8D1B55BEAB4DDBC97C632976457525A70F3CF6897E2646CF5F77FCD2590E9`。

执行记录中的 `ui_conclusion` 故意保留“待截图检查”的自动化阶段状态；上表为之后逐张实际查看 `selected.png`、`result.png` 得出的人工视觉结论。
