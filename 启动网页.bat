@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title 伽师新梅网站 - 本地预览

echo ============================================
echo   伽师新梅网站 本地预览
echo ============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [错误] 没有找到 Node.js。
  echo 请先安装 Node.js：https://nodejs.org/zh-cn  安装后重开本窗口。
  echo.
  pause
  exit /b 1
)

if not exist "node_modules" (
  echo 首次运行，正在安装依赖，大约需要 1-3 分钟，请耐心等待...
  echo.
  call npm install
  if errorlevel 1 (
    echo.
    echo [错误] 依赖安装失败。请把上面的报错整段截图发给 AI。
    echo.
    pause
    exit /b 1
  )
  echo.
  echo 依赖安装完成。
  echo.
)

echo 正在启动本地服务器，浏览器会自动打开 http://localhost:4321
echo 关闭本窗口即可停止服务器。
echo.
call npm run dev -- --open

echo.
echo 服务器已停止。
pause
