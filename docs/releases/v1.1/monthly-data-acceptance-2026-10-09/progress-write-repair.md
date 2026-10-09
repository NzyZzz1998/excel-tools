# Excel 验收进度写入修复

2026-10-09，首次 Excel 客户端验收于约 7 秒失败，尚未打开源工作簿；这是验收脚本进度文件的第二次写入失败，不是产品输出或 Excel 读取失败。原 `excel-client/` 报告和脚本副本 `validate_excel_chunked-attempt-1.ps1` 完整保留，旧脚本 SHA-256 为 `7ad4c163f2d454a6b105654375801198f18b86b024c0f9360dab6214d651f10e`。当次源/输出哈希均未变，独立 Excel 实例正常退出。

同一个 bundled pwsh 中，用两个临时合成文本文件最小复现 `[IO.File]::Replace(source,destination,$null)`：外层 `System.Management.Automation.MethodInvocationException`，HResult `-2146233087`；内层 `System.ArgumentException`，HResult `-2147024809`，固定诊断为 `The path is empty. (Parameter 'path')`。PowerShell 方法绑定中的 null 备份路径导致失败。

改为同目录 `[IO.File]::Move(source,destination,$true)` 替换。`test_progress_replacement.ps1` 从当前正式脚本 AST 提取并运行真实 `Write-ProgressRecord` 函数：创建一次、连续替换两次，验证最终身份/阶段/累计格数及无临时残片，通过。没有启动 Excel 或读取业务数据。原比较器 60,012 格合成自测重复通过。

实际失败及清理报告增加异常链的类型和 HResult，仍不记录异常消息中的业务内容。随后在全新 `excel-client-attempt-2/` 目录运行带保护的客户端校验；没有重跑产品，也没有覆盖首次失败证据。结果以该目录报告为准。

修复后脚本 SHA-256：`f583916d702a12040c154edeb0b853a2d49c7060356f2e3ed0e7d967048ae974`。第二次实际客户端验收通过：Excel 16.0.17932.21000，主表 152 块共 16,667,992 格、次表 1 块 2 格，合计 16,667,994 格，0 差异。独立监控总耗时 31.703 秒，Excel 峰值 RSS 659.133 MiB，采样 private 峰值 609.211 MiB，未触发资源/时间保护。源与输出哈希未变，只读、禁宏和外链、保存 0 次、清理错误 0，独立 Excel 正常退出。以上是客户端核验耗时及内存，不是产品转换性能。
