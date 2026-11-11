@echo off
REM Mock exam score analysis - run locally (double-click)
REM Requires: Python 3 (python.org)
echo Installing dependencies...
py -m pip install streamlit pymupdf openpyxl
if errorlevel 1 (
  python -m pip install streamlit pymupdf openpyxl
)
echo Starting the app - your browser will open automatically.
py -m streamlit run app.py
if errorlevel 1 (
  python -m streamlit run app.py
)
pause
