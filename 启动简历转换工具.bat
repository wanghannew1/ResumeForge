@echo off
rem 简历转换工具启动脚本：优先无窗口启动(pythonw)，找不到则退回 python
cd /d "%~dp0"
set "SCRIPT=%~dp0简历转换工具-解决附件图片后缀重复问题.py"
set "LAUNCH="
where pythonw >nul 2>nul && set "LAUNCH=pythonw"
if not defined LAUNCH where python >nul 2>nul && set "LAUNCH=python"
if defined LAUNCH (
    start "简历转换工具" %LAUNCH% "%SCRIPT%"
) else (
    echo [错误] 未找到 Python，请安装 Python 并勾选 Add to PATH 后重试。
    pause
)
