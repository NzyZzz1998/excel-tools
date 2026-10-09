# v1.1.1 独立输出发布验收

当前状态：**冻结源码独立验收通过。** 引擎 SHA256 为 `602f867dfd9be498e911a6599cd8da1ffcdaed5dd6bb20db5af9637b0dc196f3`。最终执行记录见 [publication-result-attempt-3.json](publication-result-attempt-3.json)，代码审查、两次探针适配及验证边界见 [publication-review.md](publication-review.md)。该结论针对源码故障路径，不代表本探针运行了打包 EXE 或断电测试。

## 契约与场景

| 场景 | 真实操作／注入位置 | 必须满足的断言 |
|---|---|---|
| 写 `.part` 时强制终止 | 自建隐藏 Python 子进程；ZIP 首成员已写入并 flush 后，以随机 nonce、PID、实际 `.part` 路径同步。父进程使用持有的 Popen 进程句柄终止，只操作自己创建的对象 | 写入期和强制终止后均无新最终 `.xlsx`；源和旧有效结果 SHA 不变；遗留 `.part` 不冒充正式结果 |
| 强制终止后正常重试 | 同一冻结引擎正常再次处理该小型合成源 | 正式结果使用可用的 `_2`；只有一个新正式文件；ZIP CRC、单元格内容／类型、合并结果及非目标部件通过独立核对；源和旧结果 SHA 不变 |
| 写盘失败叠加真实文件锁 | 在第二个 `ZipFile.open(..., 'w')` 位置，持有实际 Win32 只读句柄、允许读写但禁止删除；注入 ENOSPC | 原始 `OSError` 类型和 errno 28 保留；`partial_path` 指向本次 `.part`；`cleanup_error` 保留实际权限／占用异常；外层错误文本能看到原始原因和残片路径；不得创建最终 `.xlsx` |
| 锁失败后恢复重试 | 仅释放本探针持有的读句柄，再正常处理 | 正式结果可用、CRC和内容正确；源及旧结果不变；此前残片不被静默当成成功结果，也不被擅自删除 |

临时文件预计以同目录隐藏名称 `.<源stem>_拆分填充-<随机>.part` 写入，完整关闭 ZIP 后才发布正式名称。探针通过实际目录差集识别自己创建的唯一 `.part`，不依赖文件对象 `.name` 一定是路径（`os.fdopen` 可能显示文件描述符）。

## 来源与隔离

脚本：[probe_output_publication.py](probe_output_publication.py)。它自行生成一个只有五个预期输出格的小型 OOXML 包，不导入应用测试辅助函数；值均显式以 `SYNTHETIC` 标记。运行前强制校验给定引擎 SHA，复制到唯一 ignored `testfile/v1.1.1-publication-*` 目录，再由所有进程只导入这份冻结副本；最终复核工作区源 SHA 仍一致。报告记录脚本、引擎及结果哈希。

脚本不操作已有 PID、不扫描并终止用户程序、不启动 EXE 或 Office，不使用业务文件。子进程使用 `CREATE_NO_WINDOW`；父进程最多等待 20 秒检查点，失败也仅清理自己持有的进程对象。所有合成工作簿、冻结源码及残片保留在本次唯一 ignored 测试目录；旧证据报告拒绝覆盖。

复现命令（报告路径须选尚不存在的新文件，以保留证据）：

```powershell
python docs/releases/v1.1.1/acceptance/probe_output_publication.py `
  --engine excel_unmerge_fill.py `
  --expected-engine-sha 602f867dfd9be498e911a6599cd8da1ffcdaed5dd6bb20db5af9637b0dc196f3 `
  --report docs/releases/v1.1.1/acceptance/publication-result-new.json
```

真实原生 Tk 布局、状态截图和 GUI 错误呈现由主代理及 GUI 代理验证；本轨道不重复。本探针未重跑 104MB 全表，也不替代业务内容验收。
