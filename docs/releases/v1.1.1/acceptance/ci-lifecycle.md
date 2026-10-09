# CI 中 Tk 测试对象生命周期修正

首轮 [CI 37955040642](https://github.com/NzyZzz1998/excel-tools/actions/runs/37955040642) 在源码测试阶段退出，未到打包阶段。以下修正仅涉及测试夹具；产品源码、已验收 EXE/ZIP 和月表内容证据未改变。最终关闭该 CI 故障仍需新的远端运行结果。

## 观察与诊断界限

CI 的 `test_resume_shows_old_task_then_starts_preserved_new_selection` 附近出现两次 `tkinter.Variable.__del__` 的 `RuntimeError: main thread is not in main loop`，随后为 `Tcl_AsyncDelete: async handler deleted by the wrong thread`。这是 Tcl/Tk 资源在错误线程释放的直接证据；测试名称只是当时正在运行的用例，不能据此认定恢复业务逻辑出错。

旧 `WindowTests.tearDown` 只调用 `root.destroy()` 和临时目录清理，仍保留 `case.app`、`case.root`。销毁控件不会同时清空 Python 对象引用。测试进程会连续创建/销毁14个 Tk 窗口，并继续启动后续批次线程，因此已销毁解释器及变量不应留到后续线程可能触发的垃圾回收。

[独立弱引用探针](probe_tk_lifecycle.py) 与 [结果](ci-lifecycle-probe.json) 确认：旧清理返回时 app、root、status、all_merges 四者仍被测试对象持有；新清理返回时四者均已释放；记录的析构回调全部发生在 MainThread。探针仅创建自己的透明 Tk 窗口，基线也在主线程作最终清理。

本机尝试了后续 worker 主动 `gc.collect()`、以及完整115项测试配合每5ms后台GC压力；均未复现首轮CI的随机触发时序。因而**没有证明具体是哪一个引用循环触发了该次CI失败**，也没有把假设中的 mock/异常 traceback 循环写成确定根因。原始本地诊断日志留在忽略的 `testfile/v1.1.1-tk-gc-*.txt`。

## 最小修正及产品边界

`tests/test_windows_gui_helpers.py` 的 `tearDown` 现在确认当前线程为主线程，先销毁控件，再清空 `self.app/self.root`，在该线程执行 `gc.collect()`，并断言 app/root 的弱引用均已失效。每个真实 Tk 用例都执行此断言，而不依赖某一次随机后台GC是否命中。已有主动关闭窗口并置 `self.root=None` 的测试也由 `app.root` 弱引用覆盖。

没有新增 sleep、跳过测试、关闭GC或修改应用异常处理。产品入口每进程创建一个 root；worker 参数只有文件、选项、队列和 Event，不持有 Tk 对象；安全关闭等待当前文件完成，销毁后不创建下一测试窗口。源码审查未发现与本次测试夹具重复创建/销毁相同的产品路径。这不替代远端复验，也不外推第三方嵌入、多解释器或强制终止场景。

## 验证与身份

- [GUI及批次21/21通过](ci-lifecycle-gui-tests.txt)，包括全部14个真实Tk用例的释放断言。
- [完整115/115通过](ci-lifecycle-all-tests.txt)，本机Python3.12.8；首次失败CI为Python3.12.10，后续远端结果另记。
- `tests/test_windows_gui_helpers.py` SHA256：`4903ddaa44b454c2c7aeadb88466e64e4f48a85148a10ce65eca5ec1b77a254e`。
- GUI SHA256仍为 `50ea03ad60ed6f86003132b9d4c951213f5f01b6c202a677bf765a8df01d2094`；引擎仍为 `602f867dfd9be498e911a6599cd8da1ffcdaed5dd6bb20db5af9637b0dc196f3`。

本轮没有重跑月表、构建或操作用户正在使用的程序。清理断言用于防止测试对象跨用例残留；远端CI是否恢复通过由重跑结果决定。

## 远端复验关闭

`37956290587` 已完整通过所有GUI用例，剩余失败是独立的临时目录长短名称断言，见[路径别名修正](ci-path-alias.md)。最终[CI 37956614673](https://github.com/NzyZzz1998/excel-tools/actions/runs/37956614673)的源码测试、打包与便携自检全部通过，故本项已关闭。上文保留故障诊断时的观察边界。
