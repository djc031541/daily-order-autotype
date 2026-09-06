#!/bin/zsh
# 더블클릭하면 브라우저에서 초안 생성기가 열립니다.
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  python3 -m venv .venv
  .venv/bin/pip install -q -r requirements.txt
fi
.venv/bin/streamlit run streamlit_app.py
