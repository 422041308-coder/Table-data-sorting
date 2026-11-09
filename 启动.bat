@echo off
setlocal
cd /d "%~dp0"

rem 依次尝试：本机 DSH 自带 Python → py -3 → python，并验证能真正执行
set "PYEXE="
if exist "C:\Python312\python.exe" set "PYEXE=C:\Python312\python.exe"
if not defined PYEXE (
  where py >nul 2>nul && set "PYEXE=py -3"
)
if not defined PYEXE (
  where python >nul 2>nul && set "PYEXE=python"
)
if not defined PYEXE (
  echo 未找到 Python，请安装 Python 3.10+ 后重试。
  pause
  exit /b 1
)
%PYEXE% -c "import sys" >nul 2>nul
if errorlevel 1 (
  echo 当前 Python 不可用（可能是 Windows 应用商店占位程序），请安装 Python 3.10+ 后重试。
  pause
  exit /b 1
)

echo 正在启动「表格检索 / 模板改造 / 数据核对」...
echo 服务地址：http://127.0.0.1:8765/   （关闭本窗口即停止）
echo.
%PYEXE% -m app.server --port 8765
if errorlevel 1 (
  echo.
  echo 启动失败，请检查上方错误信息（常见原因：端口被占用 / 缺少 app 目录）。
  pause
)
