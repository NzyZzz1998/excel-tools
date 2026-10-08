@echo off
chcp 65001 >nul
setlocal DisableDelayedExpansion
set "PYTHONUTF8=1"

if not exist "%~dp0excel_unmerge_fill.py" goto missing_script

py -3 -c "import sys; sys.exit(sys.version_info < (3, 8))" >nul 2>&1
if not errorlevel 1 goto run_py

python -c "import sys; sys.exit(sys.version_info < (3, 8))" >nul 2>&1
if not errorlevel 1 goto run_python

echo 未找到 Python 3.8 或以上版本。
echo 请安装 Python 3，并保留 Tcl/Tk 文件选择组件。
echo 安装时启用 Python Launcher 或勾选 Add Python to PATH，然后重新打开本文件。
set "TASK_EXIT_CODE=1"
goto finish

:missing_script
echo 找不到 excel_unmerge_fill.py。
echo 请先解压整个工具包，并将本启动文件与 Python 脚本放在同一文件夹。
set "TASK_EXIT_CODE=1"
goto finish

:run_py
py -3 "%~dp0excel_unmerge_fill.py" %*
set "TASK_EXIT_CODE=%errorlevel%"
goto finish

:run_python
python "%~dp0excel_unmerge_fill.py" %*
set "TASK_EXIT_CODE=%errorlevel%"

:finish
echo.
pause
exit /b %TASK_EXIT_CODE%
