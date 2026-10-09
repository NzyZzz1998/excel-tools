# v1.1 进度

2026-10-09：大文件修复与验收通过，已推送并正式发布 [v1.1](https://github.com/NzyZzz1998/excel-tools/releases/tag/v1.1)，设为 Latest。当前事实源：[大文件与最终候选验收](large-data-acceptance-2026-10-09/report.md)。交付模式：Vibe Coding。

- [x] 文本、namespace、xml:space、特殊坐标、范围声明保真与不支持内容拒绝。
- [x] 大表逐行处理与临时文件承接；真实 190,566 行日明细约 63.5 秒完成，工作集峰值 71.92 MiB。双大表派生负载约 130 秒、71.98 MiB，不冒充真实月报。
- [x] 线程分配／启动内存错误恢复、停止／继续／仅失败重试，以及中文恢复提示。
- [x] 最终快照 95 项回归、9 项独立验收，32 项 DOM/stream 对照与资源清理检查通过。
- [x] 新免安装 EXE 实际处理 6 份业务文件 × 两种模式；12 份结果共 5,360,596 个实体格独立核对通过。
- [x] 两组结果重跑均无需处理 6/6；实际窗口显示 v1.1，正常退出，无残留进程。真实日明细两种结果经本机 Excel 各 2,667,930 格位对照通过。
- [x] 提交 `0678344` 已推送，远端 CI 37874774963 成功；v1.1 tag与Release绑定该提交，已验收附件上传并从正式Latest入口下载复核通过。[发布记录](large-data-acceptance-2026-10-09/release/report.md)。

当前 ZIP：`dist/ExcelTools-v1.1-Windows-x64.zip`，SHA256 `07b801517060896c01ea9be2aaf1420784d2e87fe3a5f2ae4ea1b31f5b0488f6`。先前 `92870bdc…` 候选已归档，不能作为本版下载附件。

追加真实月度文件验收：用户已提供104MB、757,636行／22列文件，已发布v1.1默认模式约6分46秒完成，进程树峰值295MiB；独立16,667,988实体格与Excel16,667,994格位对照通过，原件未变。[月度验收](monthly-data-acceptance-2026-10-09/report.md)。程序及发布ZIP没有修改。

未验证边界：此月度文件全拆分模式、更大全月明细、WPS、干净 Windows 10/11、复杂 Office 内容和人工体验补证见 [清单](manual_verification_v1.1.md)。不宣称任意月度文件容量或速度保证。

用户“深度验收，没问题就发布”承接本项目必要修复、提交、推送、v1.1 tag、Release 和免安装附件上传。仓库保持 private，原业务文件及输出仅在忽略的 `testfile/`，二进制仅作 Release 附件。无安装器或商店发布面。

历史：[前序五份小报表](business-acceptance-2026-10-09/report.md)、[逐行改造前的深验](release-acceptance-2026-10-09/report.md)、[最初候选](verification_v1.1.md)。历史失败和未验证状态保留，当前判断以同源新证据为准。
