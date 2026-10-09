入口判断：/acceptance → 发布完成

# v1.1.1 发布记录

北京时间 **2026-10-10 00:07:56** 正式发布 [v1.1.1](https://github.com/NzyZzz1998/excel-tools/releases/tag/v1.1.1)，设为 Latest。随后按用户明确授权将[仓库](https://github.com/NzyZzz1998/excel-tools)设为 public；[匿名API回读](acceptance/release-publication.json)返回200并确认公开和正式版本状态。

## 来源与验证

- 实现提交 `1a928e7`，两次测试环境适配提交 `3e26aa9`、`96b6a6e`；后两次没有修改应用源码。
- annotated tag `v1.1.1` 指向 `96b6a6ea4ae2c47c0fe50540cb326835bb907c4c`。
- [tag源码身份](acceptance/tag-source-identity.json)确认两个应用文件与冻结构建仅有Git的LF/Windows工作区CRLF行尾差异，正文完全相同；清单中的SHA针对冻结Windows字节。直接比较Git blob与工作区原始SHA会因行尾而不同，不能当作功能改动。
- [Windows CI 37956614673](https://github.com/NzyZzz1998/excel-tools/actions/runs/37956614673)通过源码回归、Windows打包、独立路径/环境便携自检和资产上传；[记录](acceptance/ci-success.json)。此前两个失败run保留，修正过程见[验收报告](acceptance/report.md)。
- 本地115项回归、9项独立验收及异常保存测试通过。真实104MB月表约6分43秒、峰值工作集295.86MiB，输出与v1.1完整核验结果逐字节一致。

## 用户下载的附件

[ExcelTools-Windows-x64.zip](https://github.com/NzyZzz1998/excel-tools/releases/download/v1.1.1/ExcelTools-Windows-x64.zip)，**12,170,380字节**。包内仅有`ExcelTools.exe`和`使用说明.txt`。

| 对象 | SHA256 |
|---|---|
| ZIP | `7d3954b5a898c14f42aef8c8881acd8f075097d941689204c68e18faa4c45b52` |
| EXE | `b5eca2e5837698e677bdb3f6b9bf844275f734ba0ce7643275fd46871ccaa2ed` |

上传的是本地完成实际月表验收的冻结包，远端资产digest相同；正式Release重新下载后再核对ZIP/EXE哈希、成员和CRC，全部一致。下载EXE在中文/空格目录、不同cwd、去除Python/Tcl/Conda变量、PATH仅System32的环境中隐藏启动自检，退出0并记录`v1.1.1 self-test passed`。见[下载验证](acceptance/release-download/release-download-verification.json)和[自检日志](acceptance/release-download/portable-self-test.txt)。

这些证据区分了源码CI新构建与实际上传的冻结包；没有把CI产物哈希当作本地包哈希，也没有以相同源代码替代下载核验。旧v1.1包和用户运行中的旧窗口保持原状；升级需解压新包并运行其中的EXE。

## 公开与边界

公开前审查历史、当前内容和后续增量，见[公开适宜性](public-readiness.md)。真实工作簿和生成结果留在本机ignored目录；仓库只含代码、合成夹具、工程统计及工具界面证据。匿名API可读；未增加许可证或改变贡献授权。

本版为进度与可靠性修补，没有新增性能提速承诺。月表验证仅覆盖默认模式；更大/不同结构文件、WPS、干净Windows和真实断电仍需对应验证。详情以[验收报告](acceptance/report.md)为准。
