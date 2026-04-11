#!/bin/bash
cd "$(dirname "$0")"
/Users/zhangdongfang/Library/Python/3.9/bin/streamlit run app.py --server.port 8502 --server.address 0.0.0.0 --server.headless true
