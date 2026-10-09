# README 独立校对与本地预览

结论：当前 README 主张、版本区分、链接和四组布局检查通过。审查只读 README/SVG，没有修改产品源码、操作用户 EXE 或截取业务窗口。

绑定对象：

- [README](../../../README.md) SHA256 `cc557dfd94994b40a3e9fe799fcc2b3ff371f867e6758bb5f6256d66e4f68725`。
- [规则示意 SVG](../../../assets/readme/merge-example.svg) SHA256 `40d70b94d5cd12e9960143d7e971f3296e8ea3831fa13b5f4cef2873f9eb2b82`。

## 内容和链接

- 第一屏依次说明用途、文件格式/本机处理、下载入口、已发布与开发中版本，再用虚构表格解释规则。使用步骤先于开发命令，支持边界和退出后不恢复任务记录均有说明。
- GitHub API 实际查询 Latest 为正式 `v1.1`（非 draft、非 prerelease），附件 `ExcelTools-Windows-x64.zip`，12,173,951 字节。README 的下载 URL 与该附件名称吻合，且明确 v1.1.1 仍在推进；没有将本地候选当作已发布包。此次只核对 Release 元数据，没有重复下载附件或运行旧包。
- 发现并反馈一处测量口径混用，主任务已修正：日明细约63.5秒/72MiB来自最终源码引擎单文件测量；月表约6分46秒/295MiB来自已发布 EXE。当前正文明确区分两者。[日明细报告](../v1.1/large-data-acceptance-2026-10-09/report.md) 与 [月表报告](../v1.1/monthly-data-acceptance-2026-10-09/report.md) 支持相应数字、规模和月表16,667,988个实体格核验。
- 六个唯一仓库相对目标均存在：使用说明、v1.1.1进度、SVG、两份历史大表报告、开发说明。Windows/Excel/WPS的已测与未测范围没有被写成全面兼容承诺。
- `beautify-github-readme/scripts/audit_readme.py` 检查通过（1个本地图像）。SVG无脚本或 `foreignObject`，示例内容明确标为虚构。

## 渲染

以 `gh api markdown` 获得当前 README 的 GitHub GFM HTML，再用本地 GitHub 风格 CSS、Playwright 和 Edge headless 渲染。HTML 的 `<img width="100%">` 嵌入 SVG，未直接导航到固定宽1200的SVG。本地 CSS预览不等同于已公开GitHub完整页面的像素复刻。

| 视口/主题 | 页面实际宽度 | SVG显示宽度 | 结果 |
|---|---:|---:|---|
| 900px / light | 900px | 852px | 无页面横向溢出、无缺图/页面错误 |
| 900px / dark | 900px | 852px | 同上，图内固定浅底与文字对比正常 |
| 360px / light | 360px | 328px | 无页面横向溢出，正文/下载入口可读 |
| 360px / dark | 360px | 328px | 同上，版本提示与示例可辨认 |

逐张查看了四组全页及首屏截图，检查章节衔接、表格、代码和最终反馈说明。手机下SVG标签约9.8px，较小；紧邻正文和替代文本已重复完整规则，不要求用户依赖小字理解操作，作为非阻断限制保留。

预览HTML、脚本、8张截图和检查JSON仅留在已忽略的 `testfile/readme-preview-v1.1.1/`，不随提交；它们只有README及虚构示例，没有业务单元格或其他应用截图。本报告为可提交结论。以后README内容或Latest版本改变，需要重新核对对应主张和入口。

## GUI 日志归档

已将被 `*.log` 忽略的两份原始日志逐字节复制为可提交文本，并修正 [GUI验证报告](acceptance/gui/review.md) 的相对引用：

- [21项GUI回归](acceptance/gui/gui-tests.txt)，SHA256 `683f98883f3794fe5b4c567063cbf5bc28ab6c6c6f2f70a0574ab919f0e31aac`，与原 `.log` 相同。
- [源码自检](acceptance/gui/source-self-test.txt)，SHA256 `0c4f279e8ab7c3d3ed4e057d76b5d18982e787b9f7361e6c0c2e53af33020908`，与原 `.log` 相同。

两份 `.txt` 未被Git忽略。此操作没有重跑或改写既有测试结论。
