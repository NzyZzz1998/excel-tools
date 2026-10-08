# Excel 数据处理工具

将 Excel 纵向合并的单元格按行拆开，并把原值填充到每一行。支持 `.xlsx`、`.xlsm`，可批量处理。所有数据在使用者电脑上处理，不上传 Excel 文件。

## Windows 免安装版本

仓库提供 Windows 自动构建配置。在 **Actions → Build Windows app** 中打开成功的构建，下载 **ExcelUnmergeFill-Windows-x64** 产物并解压，双击 `ExcelUnmergeFill.exe`。

这个 `.exe` 包含所需运行环境，使用者不需要安装 Python，也不需要部署服务。界面可以选择多个 Excel 文件，并查看每个文件的处理结果。输出保存在原 Excel 所在目录，自动增加 `_拆分填充` 后缀；已有结果自动加序号，不覆盖原文件。

只有构建成功且产物自检通过后，才有可下载的 `.exe`。源码 ZIP 和之前的 `.bat` 版本仍需要 Python，不能当作免安装版。

## 处理规则

- `A2:A5` 合并显示“华东仓”，拆分后 A2、A3、A4、A5 都填入“华东仓”。
- 默认保留仅横向合并的表头。跨行又跨列的 `A2:B4` 拆成 `A2:B2`、`A3:B3`、`A4:B4`，每行显示原值。
- 勾选“同时拆分横向合并”后，所有合并区域都拆开并填满。
- 普通空白保持原样；原值是零、布尔值、日期或带前导零的文本时，保留原始类型及数字格式。
- 默认处理全部工作表，包括隐藏工作表。
- 合并区域源单元格含公式时，提示先在 Excel 副本中“粘贴为值”；其他位置的公式保持原样，不计算公式。
- 合并区域内若仍有被掩盖的独立内容，停止该文件，避免静默覆盖。

## 源码运行

需要 Python 3.8 或以上，核心脚本无需第三方库。图形界面还需要 Python 自带的 Tcl/Tk 组件。

```bash
python excel_unmerge_gui.py
python excel_unmerge_fill.py "原始数据.xlsx"
python excel_unmerge_fill.py "原始数据.xlsx" --all-merges
python excel_unmerge_fill.py "原始数据.xlsx" --sheet "明细"
```

Windows 也可以双击 `Windows双击运行.bat`，Mac 可以双击 `双击运行.command`。这些源码入口要求本地已有 Python。

## 验证与构建

测试使用代码生成的合成工作簿，不需要任何真实业务数据。安装测试依赖后运行：

```bash
python -m pip install openpyxl
python -m unittest discover -s tests -v
```

GitHub Actions 在 Windows 环境执行测试、用 PyInstaller 打包，然后运行打包后的程序自检。源码测试、Windows 产物自检和实际文件的桌面验收是不同验证范围；构建成功不表示已经核对所有业务文件。

只修改有关工作表 XML，其他 ZIP 条目原样复制。新增格使用源单元格格式，拆分后边框外观可能与原合并块不同。旧版 `.xls`、密码加密及带文档数字签名的工作簿不在支持范围内；`.xls` 请先另存为 `.xlsx`。

更详细的源码使用方法见 [使用说明](使用说明.md)。
