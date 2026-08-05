#!/bin/bash
cd /home/site/wwwroot
/home/site/wwwroot/antenv/bin/streamlit run streamlit_app.py \
  --server.address 0.0.0.0 \
  --server.port $WEBSITES_PORT \
  --server.headless true
