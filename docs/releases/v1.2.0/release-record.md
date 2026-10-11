入口判断：/acceptance → 发布完成

交付模式：Vibe Coding 全流程型。

# v1.2.0 发布记录

北京时间 **2026-10-11 22:00:32** 正式发布 [v1.2.0](https://github.com/NzyZzz1998/excel-tools/releases/tag/v1.2.0)并设为Latest，仓库保持公开。用户“没问题的话推送1.2.0吧”的授权已完成；匿名API确认非draft、非prerelease，见[发布回读](acceptance/release-publication.json)。

## 源码、CI与继承证据

- 实现提交`ed85c03499fc74ad177b42b2ba443439287dfe86`；测试兼容修正`0f56c96e7f89bae22759055457f57ed6d1d5172a`为annotated tag `v1.2.0`的目标。
- 首轮CI 38145129384仅一处临时目录长短路径文本比较失败。修正为目录数量＋samefile身份，保留错误/源与历史结果断言；定向28项通过。[首轮记录](acceptance/ci-attempt-1.json)保留。
- [Windows CI 38145277446](https://github.com/NzyZzz1998/excel-tools/actions/runs/38145277446)已通过全套源码测试、打包、隔离便携自检与资产生成，见[CI结果](acceptance/ci-success.json)。
- 所有5个运行模块与实际验收的冻结源码内容一致；Git的LF与本地CRLF可能影响原始字节哈希，已逐文件核对正规化内容，见[源码身份](acceptance/release-source-identity.json)。测试修正未改应用，也没有重打EXE。
- 继承180项本地回归、真实757635行月表三模式全量核对、共享字符串压力、真实Tk和Excel客户端证据，详见[验收报告](acceptance/report.md)。没有把CI生成的另一个EXE当作实际月表验收包。

## 正式免安装附件

[ExcelTools-Windows-x64.zip](https://github.com/NzyZzz1998/excel-tools/releases/download/v1.2.0/ExcelTools-Windows-x64.zip)，**13,145,189字节**，仅含`ExcelTools.exe`及`使用说明.txt`。

| 对象 | SHA256 |
|---|---|
| 正式ZIP | `d60c15554317dfed91822115df62e7b2d072c6dcbd48fd219ea0af8bacf8cce6` |
| EXE | `51ff57528044e46a91bb3779b26a3b9d0a9feff17affe42e01fc3cda8a67347a` |

正式ZIP将说明中的候选状态改为正式版，因此与原候选ZIP不同；EXE逐字节保持已验收版本。[发布清单](acceptance/release-package.json)及[本地正式包复验](acceptance/release-package-check/release-download-verification.json)记录这一差异，原候选文件未覆盖。

从公共`releases/latest/download/ExcelTools-Windows-x64.zip`重新下载后，ZIP/EXE哈希、尺寸、成员和CRC均与发布清单一致。解压至独立中文/空格路径，在不同工作目录、移除Python/Tcl等开发变量、PATH仅System32的环境隐藏运行自检，退出0并记录`ExcelTools v1.2.0 self-test passed.`。见[下载核验](acceptance/release-download/release-download-verification.json)与[自检日志](acceptance/release-download/portable-self-test.txt)。

本机正式包位于`dist/ExcelTools-v1.2.0-release/`，正式重新下载的文件位于`testfile/release-v1.2.0-downloaded-20261011/`。旧窗口和旧快捷方式不会被自动替换；升级需完整解压新包并运行其中EXE，核对窗口v1.2.0。

## 边界与回退

新增工具为数据型输出；公式需先粘贴为值，参与区域合并需先展开。源文件与历史结果不覆盖。月表性能不含前置整理及核验，不能承诺其他文件/电脑同速。WPS、其他DPI、干净Windows和更大/不同结构文件仍需对应验证。

本次发布源码与便携包，无安装器或商店包，无配置迁移。业务原件和结果留在ignored目录，增量公开审查见[公开适宜性](public-readiness.md)。v1.1.1保留供回退；收尾文档另行提交，不改变tag与已验收资产。
