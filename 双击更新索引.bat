@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul && (py -3 sync_docs_index.py --open & goto :end)
where python >nul 2>nul && (python sync_docs_index.py --open & goto :end)
python3 sync_docs_index.py --open
:end
echo.
pause
