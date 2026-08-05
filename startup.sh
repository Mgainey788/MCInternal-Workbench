#!/usr/bin/env bash
set -euo pipefail

export PYTHONUNBUFFERED=1
PORT_NUMBER="${PORT:-${WEBSITES_PORT:-8000}}"

exec python3 -m streamlit run streamlit_app.py --server.port "${PORT_NUMBER}" --server.address 0.0.0.0 --server.headless true
