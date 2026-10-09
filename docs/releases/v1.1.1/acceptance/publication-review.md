# v1.1.1 输出发布独立验收

结论：Windows 当前冻结实现未发现输出发布阻断项。独立小型合成探针通过，源文件及已有有效结果保持不变；写入中强制终止不产生最终命名的损坏工作簿。

被测引擎 SHA256：`602f867dfd9be498e911a6599cd8da1ffcdaed5dd6bb20db5af9637b0dc196f3`。执行时复制到唯一 ignored 临时目录，所有探针只加载该副本，完成后复核工作区引擎未变。最终机器可读证据：[publication-result-attempt-3.json](publication-result-attempt-3.json)。方法与脚本：[plan.md](plan.md)、[probe_output_publication.py](probe_output_publication.py)。

## 独立证据

| 场景 | 实际结果 |
|---|---|
| ZIP 首成员开始写入后强制终止自建隐藏进程 | 检查点为 113 字节 `.part`，写入期及终止后没有新正式 `.xlsx`；随机 nonce、实际 PID、引擎哈希及临时文件路径均在终止前匹配 |
| 强制终止后正常重试 | 生成唯一有效 `_2.xlsx`；ZIP CRC、5 个预期单元格的值／类型、合并结果、非目标部件均通过；原源和旧有效结果 SHA 不变；此前 `.part` 不被静默删除 |
| 真实 Win32 读句柄禁止删除，并注入 ENOSPC | 对外仍为同一 `OSError`、errno 28；实际删除失败为 `PermissionError` / WinError 32；`partial_path`、`cleanup_error` 及可见错误文本保留原始原因和残片位置；无新最终文件 |
| 释放本探针读句柄后重试 | 新正式结果内容及 CRC 正确；源、旧结果及之前残片的 SHA 不变 |

没有操作用户已有进程、EXE、Office 或业务工作簿。所有合成文件、残片和冻结副本留在报告记录的 ignored `testfile/v1.1.1-publication-*` 子目录。

## 代码审查

- [`process_file`](../../../../excel_unmerge_fill.py#L726) 在同目录创建唯一 `.part`；所有成员及 ZIP 中央目录完整关闭后才调用发布函数（780–808 行）。失败清理限于本次 `.part`，没有删除历史结果或源文件的路径。
- [`_publish_no_replace`](../../../../excel_unmerge_fill.py#L673) 的 Windows `os.rename` 不覆盖已存在目标；[`_publish_part`](../../../../excel_unmerge_fill.py#L698) 只在确认目标名称占用时递增编号。真正发布权限失败会向上传播，不因误判占用无限寻找新名称。
- [`_cleanup_part`](../../../../excel_unmerge_fill.py#L663) 清理失败保留原异常身份及 errno，附上实际清理异常和残片路径；真实锁叠加 ENOSPC 已独立验证该分支。
- 非 Windows 分支以排他硬链接发布，并单独记录已发布但临时名称删除失败的状态；本机没有验证该分支或不支持硬链接的文件系统，不能将此 Windows 结果扩展为跨平台实测结论。

## 探针适配记录与边界

保留的 [第一次结果](publication-result.json) 使用旧版 `shutil.copyfileobj` 拦截点，候选已改为分块复制，因此没有命中检查点；是探针适配失败。[第二次结果](publication-result-attempt-2.json) 已命中 `.part`，但 Windows venv 启动器的 PID 与实际 Python PID 不同，所有权断言提前失败。最终改为直接启动同一运行时的 `sys._base_executable`，PID 一致后才实施强制终止。前两次没有被计为产品验收通过。

本结果验证进程被终止时的命名隔离及普通异常下的可恢复性，不验证机器断电、磁盘控制器缓存或网络文件系统持久化；实现没有新增 `fsync` 持久化承诺。未以全盘填满方式触发 ENOSPC，而是在真实写入位置注入该错误，并使用真实 Windows 删除共享限制验证清理失败。未重复既有全量内容测试。
