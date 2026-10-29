@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (py rakuten_collect.py) else (python rakuten_collect.py)
pause
