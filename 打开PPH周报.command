#!/bin/bash

set -u

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
PORT="3001"
URL="http://localhost:${PORT}/"
LOG_DIR="${PROJECT_DIR}/.logs"
LOG_FILE="${LOG_DIR}/pph-web.log"

export PATH="/usr/local/bin:/opt/homebrew/bin:${HOME}/.nvm/current/bin:${PATH}"
cd "${PROJECT_DIR}" || exit 1

page_is_ready() {
  /usr/bin/curl -fsS --max-time 2 "${URL}" 2>/dev/null \
    | /usr/bin/grep -q "PPH周报系统"
}

if ! page_is_ready; then
  /bin/mkdir -p "${LOG_DIR}"

  if ! command -v npm >/dev/null 2>&1; then
    echo "未找到 npm，请确认 Node.js 已安装。"
    echo "按回车键关闭窗口。"
    read -r
    exit 1
  fi

  echo "正在启动 PPH 周报系统……"
  nohup npm run dev -- --port "${PORT}" >"${LOG_FILE}" 2>&1 &
  echo $! >"${PROJECT_DIR}/.pph-web.pid"

  for _ in {1..60}; do
    if page_is_ready; then
      break
    fi
    /bin/sleep 1
  done
fi

if page_is_ready; then
  /usr/bin/open "${URL}"
  echo "PPH 周报系统已打开。"
  /bin/sleep 1
  exit 0
fi

echo "启动超时，请查看日志：${LOG_FILE}"
echo "按回车键关闭窗口。"
read -r
exit 1
