@echo off
title WCG Panel Takeoff
cd /d "%~dp0"
python -m streamlit run app.py --server.port 8501 --browser.gatherUsageStats false
pause