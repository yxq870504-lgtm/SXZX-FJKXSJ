@echo off
setlocal
chcp 65001 >nul
title 放假开学时间自动化看板服务
cd /d C:\Users\yangxq01\lobsterai\project
where python >nul 2>nul
if errorlevel 1 (
  echo [错误] 未找到 python，请先安装 Python 或把 Python 加入 PATH。
  pause
  exit /b 1
)
where git >nul 2>nul
if errorlevel 1 (
  echo [错误] 未找到 git，请先安装 Git 或把 Git 加入 PATH。
  pause
  exit /b 1
)
echo 正在启动本地看板服务...
echo 本地访问地址：http://127.0.0.1:8788/index.html
echo GitHub Pages：https://yxq870504-lgtm.github.io/SXZX-FJKXSJ/
echo.
python dashboard_server.py
if errorlevel 1 (
  echo.
  echo [错误] 服务异常退出，请查看上方日志。
)
pause
