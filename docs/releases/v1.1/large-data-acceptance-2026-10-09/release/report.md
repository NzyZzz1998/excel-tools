入口判断：/acceptance → v1.1 发布闭环

交付模式：Vibe Coding。**结论：通过，已发布。** 仅针对[总报告](../report.md)明确的业务与环境范围，真实100MB月度表仍待验证。

- 源码提交／tag：`067834455de167d86351b5ab338c518744ec4afa`，已推送origin/main，远端v1.1指向同一提交。[冻结源码与提交对应](source-commit-binding.json)逐项核对2个程序、5个测试和包内说明；仅Git CRLF正规化差异。构建时旧HEAD加未提交内容的血缘已闭合。
- 远端CI：[37874774963](https://github.com/NzyZzz1998/excel-tools/actions/runs/37874774963)成功。实际日志显示95项测试通过，构建、清理Python路径后的EXE自检、包整理及上传均成功。[状态](ci-run.json)、[原始日志](ci-steps.txt)。CI产物独立构建；Release使用此前实际业务验收的冻结本地包，没有用CI重新生成包替换其身份。
- 正式发布：[v1.1](https://github.com/NzyZzz1998/excel-tools/releases/tag/v1.1)，2026-10-09T02:30:59Z，Latest=true，draft=false，prerelease=false。仓库private保持。[Release身份与资产](published-release.json)。
- 唯一资产：`ExcelTools-Windows-x64.zip`，12,173,951字节，SHA256 `07b801517060896c01ea9be2aaf1420784d2e87fe3a5f2ae4ea1b31f5b0488f6`。远端asset digest、本地已验收包、下载包全部一致。EXE SHA256 `239b1456771ae417f044dfd1ac2642f7ac36de7bc3781c6d2ac3ddc75d2d7ea4`，真实GUI已确认标题v1.1。
- 发布前先下载草稿资产并在中文空格路径、独立cwd、无Python/Tcl环境、受限PATH、独立临时目录运行隐藏EXE自检，退出0、日志`ExcelTools v1.1 self-test passed.`。[自检记录](download-self-test.json)。发布后使用不带tag的`gh release download`从Latest再次下载，[完整字节对照](latest-download-integrity.json)与已自检文件相同，未用缓存路径覆盖或复用旧v1.0包。

用户原“深度验收，没问题就发布”覆盖必要的提交、推送、tag、Release与附件上传。本轮没有创建安装器或商店包，没有更改仓库可见性或上传业务数据。该目录仅含发布元数据、日志、哈希和验证脚本。

后续文档提交仅记录发布结果，不更改应用、包、tag或已验收输入；对应源码CI和本地运行证据继续有效。原始CI／失败日志保留终端尾部空白，格式检查定向排除原日志文件，源代码／文档正常检查，不改写失败证据。

升级：关闭旧版，完整解压新包到新目录并确认标题v1.1；旧快捷方式可能仍指向v1.0.0。无需数据迁移，原表保留。旧v1.0.0 Release仍可下载回退，但其已知大文件和保真问题仍存在。
