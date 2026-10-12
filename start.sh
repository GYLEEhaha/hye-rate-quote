#!/bin/bash
# 启动内部运价查询本地服务
cd "$(dirname "$0")"
PORT="${1:-8765}"
echo "环洋 HYE 运价查询 → http://127.0.0.1:${PORT}/"
echo "更新价格: 替换 data/rates.xlsx 后执行 python3 scripts/parse_rates.py"
exec python3 -m http.server "$PORT"
