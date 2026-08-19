@echo off
rem 全量单元测试一键回归（详见 tests/run_all.py）
rem 用法：跑测试.bat [--all] [--filter 关键字] [--timeout 秒]
cd /d "%~dp0"
.venv\Scripts\python.exe tests\run_all.py %*
exit /b %ERRORLEVEL%
