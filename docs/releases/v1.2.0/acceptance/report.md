入口判断：/acceptance → 本地候选验收

交付模式：Vibe Coding 全流程型。2026-10-11。当前结论：实现、便携包与发布下载验收通过，A01–A10均已闭合。v1.2.0已正式发布并设为Latest，详见[发布记录](../release-record.md)。

# 对象与证据身份

起点分支main，base HEAD `b48e97dd77bcd32c1508eab076a924568a6d54bf`；候选构建时新实现未提交，采用冻结源文件SHA256绑定。旧拆分填充引擎未改，SHA256仍为 `602f867dfd9be498e911a6599cd8da1ffcdaed5dd6bb20db5af9637b0dc196f3`。下表保留候选构建身份，见[candidate-identity.json](candidate-identity.json)。正式tag对应`0f56c96e7f89bae22759055457f57ed6d1d5172a`，所有5模块与冻结源码仅有可能的Git行尾正规化，内容一致，见[发布源码身份](release-source-identity.json)；包内说明更新后的正式ZIP见[发布包身份](release-package.json)，EXE不变。

| 身份闭包 | 状态与证据 |
|---|---|
| 候选路径/版本/哈希 | 已提供；dist/ExcelTools-v1.2.0/ExcelTools.exe，v1.2.0，SHA256 `51ff57528044e46a91bb3779b26a3b9d0a9feff17affe42e01fc3cda8a67347a` |
| source commit或等价构建身份 | 已提供；候选构建时采用base HEAD＋dirty工作树5个模块SHA及独立冻结快照；发布后另以release-source-identity绑定真实tag源码 |
| review范围/时间 | 已提供；2026-10-11新IO、编排、GUI和根窗口集成差异，由实现之外的角色做黑盒测试和交叉审查 |
| 日志/测试时间/环境 | 已提供；2026-10-11 Windows本机，Python3.12.8、Tk8.6.15、openpyxl3.1.5（仅测试）、PyInstaller6.22.3 |

自动化、源码审查与包自检依赖该闭包。月表源码引擎性能另绑定两个数据模块哈希，不冒称打包EXE进程树性能。

# 已执行验证

| ID / 合同 | 预期 | 实际与证据 | 结论 |
|---|---|---|---|
| A01 / T01–08 | 新旧行为回归 | 180项通过，15.372秒；旧115＋IO17＋编排28＋新GUI13＋根窗口7。日志testfile/v1.2.0-acceptance/full-tests-locked.log | 通过 |
| A02 / T04–06 | 类型/公式/合并/日期/文件保护 | openpyxl独立oracle；表头重排、带缓存/无缓存公式均拒绝、1904保留/混用拒绝、字面等号文本、CR、来源冲突、行数边界、无漏行。tests/test_table_io.py与test_table_tools.py | 通过 |
| A03 / T06–08 | 失败取消不发布半套，原件和旧结果不动 | Windows真实读锁、ENOSPC、发布权限错误、清理失败、源mtime改变、SQLite索引中断、ZIP打包中取消、两个并发发布与旧结果哈希断言 | 通过 |
| A04 / T01/08 | 新旧Tk互斥、状态正确、布局可用 | 真实Tk测试、后台线程启动失败、晚取消不吞成功、运行中根关闭仅请求取消；主线程weakref/gc回收通过。默认/最小界面见[本机DPI证据](gui/dpi-review.md) | 通过（本机96DPI） |
| A05 / T08 | 便携包不依赖开发路径，版本正确 | 独立中文目录、不同工作目录、移除Python/Tcl/虚拟环境变量、最小PATH＋独立TEMP；退出码0，日志“ExcelTools v1.2.0 self-test passed.”；旧处理/三种新任务/真实Tk入口都执行 | 通过 |
| A06 / T04 | Excel客户端打开新xlsx并解释类型正确 | 独立创建的Excel16.0 COM实例只读打开合成输出，3行×4列，编号0012、日期序列、布尔、数字0000格式、字面等号文本均符合；关闭工作簿并Quit自己实例。见[excel-client.json](excel-client.json) | 通过（合成样本） |
| A07 / T03/07 | 真实月表完成并保值，记录资源 | 按10万行拆8份141.138秒/30.79MiB RSS、重新合表143.070秒/30.64MiB RSS、按第1列值拆29份151.967秒/46.40MiB RSS；源哈希不变，均757635数据行。全量解码后值/类型/数字格式摘要一致；38份输出表头一致、29份分组归属正确。见[内容核验](monthly-oracle.json)、[表头与分组](monthly-groups.json) | 通过 |
| A08 / T07 | 大SST不能全表字符串驻内存 | 100000唯一字符串×544字节；SST未压缩56000113字节。拆成40000/40000/20000行，11.526秒、峰值RSS32.20MiB；独立openpyxl源与结果逐行文本/类型/摘要一致，源哈希不变。[完整证据](sst-stress.json) | 通过 |
| A09 / T08 | 远端Windows环境测试、构建与便携自检 | CI 38145277446对应正式源码commit，测试、打包、隔离EXE自检及资产生成全部通过。[CI记录](ci-success.json) | 通过 |
| A10 / 发布 | 用户下载的附件身份与版本正确 | 匿名Latest为正式v1.2.0，远端digest、从latest下载的ZIP/EXE哈希与发布清单一致；中文/空格目录、不同cwd、最小PATH、去开发环境自检退出0，日志v1.2.0。[下载证据](release-download/release-download-verification.json) | 通过 |

IO定向测试曾发现最后保存回调刚设取消但未再次检查，已补检查并按原反例通过。CR异常确认是测试夹具先把CR写成XML字面回车导致源读回先变LF，已用合法字符实体构造并先独立读回核实，没有放宽应用断言。首次全套唯一失败为旧GUI测试硬编码v1.1.1，版本合同更新后全套通过。

月表首次全量读回的原始摘要不一致，失败证据保留在[monthly-oracle-raw-failed.json](monthly-oracle-raw-failed.json)。首差F246为源inlineStr中的`_x003D_`与输出等号：openpyxl3.1.5未解释inlineStr的ST_Xstring编码，而本机Excel16只读打开真实源和输出后Value2相同、均为文本且无公式，见[真实单格证据](excel-escape-difference.json)。六种纯合成用例再验证等号、转义下划线、CR与代理对的Excel语义，见[解释及可复验探针](xstring-probe/finding.md)。这不能替代全表一致性核对。

修正仅作用于独立验收脚本：对源与输出同时进行单遍ST_Xstring解码，并先扫描所有工作表，遇到SST或缓存字符串单元格即拒绝该校验模式，防止对openpyxl已转换的SST二次解码。三个独立脚本回归测试通过，覆盖上述六例与存储类型拒绝；运行时代码和冻结候选未改。带`--inline-xstring`的全量复验实际通过，404.611秒；源、按行拆分、合回及分组均757635行，有序摘要均为`bf55620f8c01c793e07319a0eabb8240d3d86c6587039db753aea062abdda8cf`，无序行摘要也一致。

本次39本工作簿的40个worksheet部件均在校验扫描范围，源唯一SST是137字节、0条目的空表，其余输出无SST，见[存储盘点](oracle-storage-inventory.json)。这项补偿只用于已盘点的本批文件。内容摘要比较第一工作表表头后的非空行，保留文本空格，裁去行尾值为None的空单元格；日期经openpyxl解释为日期/时间。结论是值、类型、数字格式的解码后语义对照，不是原始XML逐字节或全部数字序列精度证明。按行拆分/合表核对有序摘要；分组核对行数与无序行摘要以证明整批内容守恒。额外[独立检查](verify_monthly_groups.py)耗时54.566秒：38份输出的22列表头完全一致；29份分组文件各自只有一种第1列值/类型，组间键均不同，共757635行，未跳过任何空行；全部文件读取前后size/mtime不变。核验耗时不计入处理性能。

# 月表前置与统计边界

用户原月表为104185338字节、主表757636行×22列。历史默认拆分填充结果117674593字节仍保留尾行`A757636:G757636`横向合并；直接给新工具时，本轮实际拒绝并清理整批，无正式结果。为新工具建立合法输入，在本轮忽略目录副本调用既有`process_file(..., all_merges=True)`，展开1处、填6格，205.958秒。该时间单独记录，不计入新拆表/合表成绩；不能声称原始104MB无需整理就成功。

完整展开后的输入117674551字节，SHA256 `b7c9a7dde0d6a1011b0f5ea6179592d0f5abeab39b861ae28aeb0afdaa59f3ca`。选择第一工作表、表头第1行。处理表头后全部非空行，包含汇总尾行；757635不是自动剔除汇总后的业务明细数。大表运行关闭来源附加列以便独立全量对照，来源功能另有定向测试。性能为本机Python引擎壁钟、100ms采样当前进程RSS/private，排除输入准备、哈希与oracle；不能当作任意电脑/任意结构的SLA，也不能和旧拆分填充时间直接等价比较。

脚本：[measure_tables.py](measure_tables.py)，[共享字符串探针](measure_shared_strings.py)。资源与文件哈希见[monthly-performance.json](monthly-performance.json)。本机Intel Core Ultra7 265K、20逻辑核、47.4GiB RAM、Windows11；只有这些实际样本可得到上述结论。真实源及结果全部位于忽略的testfile目录；公开证据只保留计数、哈希、资源、配置，不能提交业务行、源工作簿或分组值。

# 工具与证据来源

PowerShell/Python实际执行源码、unittest、PyInstaller和文件哈希；独立openpyxl读取类型/值/格式；psutil做单进程采样；Win32 PrintWindow仅捕获本任务创建的窗口；Excel COM仅创建并释放本任务实例。浏览器不适用于原生Tk本体。本机客户端可用，未对WPS、干净Windows10/11和其他DPI作完成声明。源码在全套测试后未改运行模块，文档与追加证据不使现有代码验证失效。

# 结论与发布面

本轮合同内无未关闭阻塞。本机96DPI、已测结构和样本的结论已闭合；其他DPI、WPS、干净Windows环境、更大或不同结构工作簿仍没有相应完成声明。

已按用户明确授权完成提交、推送、annotated tag及正式Release，Latest现为v1.2.0。正式包`dist/ExcelTools-v1.2.0-release/ExcelTools-Windows-x64.zip`仅更新包内说明，EXE与已验收候选相同；正式重新下载后再次确认哈希及自检成功。不存在新增安装器或商店包，仓库保持已公开状态。回退直接使用v1.1.1，源和历史结果保留，无配置迁移。

首轮CI 38145129384的180项中仅一处Windows临时目录长短路径文本比较失败；改为唯一目录数量断言＋samefile身份，所有错误/清理/源与旧结果哈希断言保持。定向28项通过后，新CI 38145277446全部通过。该修正仅改测试，运行模块和已验收EXE未改，既有真实月表证据仍有效；旧[失败记录](ci-attempt-1.json)保留。
