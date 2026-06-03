@echo off
chcp 65001 >nul
title 汽车之家经销商点检任务

echo ============================================================
echo   汽车之家奇瑞经销商点检任务
echo   启动前请先用记事本打开 task.py，修改顶部两个路径：
echo     dealer_list_xlsx  ^<-- 经销商名单 .xlsx
echo     standard_xlsx     ^<-- 报价标准 .xlsx
echo ============================================================
echo.

cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.11+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/2] 安装依赖...
pip install -r requirements.txt -q
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络
    pause
    exit /b 1
)

echo [2/2] 开始采集...
echo.
python task.py

echo.
echo ============================================================
echo   任务结束。输出文件在 output\ 目录。
echo ============================================================
pause
