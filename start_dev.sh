#!/bin/bash
# 测试版本启动脚本 - 端口8503
# dev分支专用，测试通过后合并到main重启正式版
cd "$(dirname "$0")"
export STREAMLIT_SERVER_PORT=8503
export STREAMLIT_SERVER_ADDRESS=0.0.0.0
python3 -m streamlit run app.py --server.port 8503 --server.address 0.0.0.0
