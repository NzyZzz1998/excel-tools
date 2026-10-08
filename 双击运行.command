#!/bin/zsh
SCRIPT_DIR="${0:A:h}"
python3 "$SCRIPT_DIR/excel_unmerge_fill.py" "$@"
RESULT=$?
echo ""
read -r "REPLY?按回车键关闭窗口……"
exit "$RESULT"
