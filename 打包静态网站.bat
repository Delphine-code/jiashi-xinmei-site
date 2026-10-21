@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title 伽师新梅网站 - 打包与预览

echo 正在构建静态网站到 dist 目录...
echo.
call npm run build
if errorlevel 1 (
  echo.
  echo [错误] 构建失败。请把上面的报错整段截图发给 AI。
  echo.
  pause
  exit /b 1
)

echo.
echo 构建成功，正在本地预览构建结果...
echo 浏览器会自动打开 http://localhost:4321
echo.
call npm run preview -- --open

pause
