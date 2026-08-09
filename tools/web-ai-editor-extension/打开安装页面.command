#!/bin/bash

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
open "$SCRIPT_DIR"
open -a "Google Chrome" "chrome://extensions/"

cat <<'EOF'
Web AI 页面修改器安装步骤：
1. 在扩展页面打开“开发者模式”
2. 点击“加载已解压的扩展程序”
3. 选择刚刚打开的 Web AI 页面修改器文件夹
4. 如果要修改本地HTML，在扩展详情中打开“允许访问文件网址”
EOF
