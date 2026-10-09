# v1.1.1 引擎实现交接

2026-10-09，依据本版 PRD F-11/F-12 与发布后 review 实现。首候选引擎 SHA-256：`602f867dfd9be498e911a6599cd8da1ffcdaed5dd6bb20db5af9637b0dc196f3`。本交接没有提交、打包或运行用户 EXE。

## 保存行为

完整 ZIP 先写源目录独占创建的随机 `.part`，名字包含截短的源 stem（避免临时名因附加随机后缀超长）。ZIP 和底层文件句柄关闭后，Windows 通过不覆盖的 rename 发布；其他平台通过不覆盖的 hard link 发布后删除本次 `.part`。文件、目录和竞争占名顺延，真实权限错误传播。没有 replace，也不清理其他会话的 `.part`。

写入/发布失败仅清理本次临时文件。若清理也失败，保留同一个原异常对象及类型/errno，附加 `partial_path`、`cleanup_error`，简短错误文本也包含残片路径。非 Windows 分支若 hard link 已成功而临时名删除失败，会明确说明完整正式结果已经发布及临时路径，不把它当作命名竞争继续生成第二份正式结果。

这解决进程被强制结束时半成品使用最终工作簿名的问题；不是物理断电或文件系统持久化保证，没有新增 fsync、恢复旧残片或途中取消。

## 进度行为

`process_file(source, all_merges=False, sheets=None, progress=None)` 保持原三参数兼容。回调收到固定五键 dict：phase、sheet、completed、total、unit。普通观察者 Exception 被隔离，不改变工作簿结果；BaseException 传播并清理本次临时文件。

- 中间事件按 monotonic clock 约 0.3 秒节流，阶段起止强制发送。
- 大表 reading 为真实已解析物理行数，总量在解析结束前未知；filling 为实际处理行数与待处理行总数，包含保留行和新增行，不等于新填充值格数。
- 小 DOM 表至少报告读取和转换起止；不伪造未知转换百分比。
- saving 为已写入 ZIP 成员的未压缩字节数/总字节数，按 1 MiB 块推进，不是压缩后文件大小。最后字节事件发生在完整 ZIP 关闭后，随后立即发布，最终成功仍以函数返回为准。
- 无匹配不创建 `.part`、不发送 saving。

XML 转换、流式阈值、公式/元数据/隐藏内容拒绝、默认横向规则未作性能重构。

## 验证

新增测试先确认旧实现没有 `.part` 写入期且不接受 progress 参数，再实现转绿。当前命令：

`python -m unittest test_engine_io_progress test_engine_regressions test_excel_unmerge test_streaming_engine -v`

在 tests 与临时 openpyxl 3.1.5 加入 PYTHONPATH 后执行：**94 项通过，1.300 秒**。组成：14 项新 I/O/进度测试、22 项原语义回归、15 项原工作簿测试、43 项强制流式回归。该结果仅为引擎相关测试，不冒称 GUI/发布候选已经通过。

新测试实际覆盖：

- 写入期只看到 `.part`，正式命名发布前可完整读取 ZIP；首个正式名称被其他写者抢占时顺延。
- 4 个线程同步竞争发布，四个不同最终名称，源及旧结果 SHA 不变，无本次残片。
- 第二个 ZIP 成员真实写入 5 字节后 ENOSPC；自己的残片清理，其他会话 `.part` 和旧结果不动。
- 再注入删除受阻时，原错误对象/errno 保留且错误消息含路径；真实 Win32 锁另由独立验收补证，不把此 mock 宣称为真实锁测试。
- 真正发布权限错误只尝试一次；POSIX hard-link 分支在本机文件系统执行同名保护及已发布/清理失败分支，但不是 Linux 整体运行认证。
- 失效 callback 与正常 callback 输出 SHA 一致；BaseException 清理；固定时钟下 1,000 原行扩至 1,200 行仅 6 个起止事件，推进时钟才产生节流中间事件；5 MiB ZIP 部件产生保存中间计数。

原 `test_output_permission_error_without_directory_is_not_retried` 的注入位置由旧 `Path.open('xb')` 改到 `mkstemp`，保留单次调用、无新增文件、源 SHA 不变断言；没有放松成功条件。复用旧独立验收脚本时须同样更新已失效的 seam，ZIP.open 真实部分写入 seam 保持有效。
