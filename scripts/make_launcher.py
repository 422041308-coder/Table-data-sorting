# -*- coding: utf-8 -*-
"""生成 GBK + CRLF 的 启动.bat（cmd 对 UTF-8/LF 的 .bat 支持很差）。"""
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

content = (
    "@echo off\r\n"
    "setlocal\r\n"
    'cd /d "%~dp0"\r\n'
    "\r\n"
    'rem 依次尝试：本机 DSH 自带 Python → py -3 → python，并验证能真正执行\r\n'
    'set "PYEXE="\r\n'
    'if exist "C:\\Users\\user\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe" '
    'set "PYEXE=C:\\Users\\user\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe"\r\n'
    "if not defined PYEXE (\r\n"
    '  where py >nul 2>nul && set "PYEXE=py -3"\r\n'
    ")\r\n"
    "if not defined PYEXE (\r\n"
    '  where python >nul 2>nul && set "PYEXE=python"\r\n'
    ")\r\n"
    "if not defined PYEXE (\r\n"
    "  echo 未找到 Python，请安装 Python 3.10+ 后重试。\r\n"
    "  pause\r\n"
    "  exit /b 1\r\n"
    ")\r\n"
    "%PYEXE% -c \"import sys\" >nul 2>nul\r\n"
    "if errorlevel 1 (\r\n"
    "  echo 当前 Python 不可用（可能是 Windows 应用商店占位程序），请安装 Python 3.10+ 后重试。\r\n"
    "  pause\r\n"
    "  exit /b 1\r\n"
    ")\r\n"
    "\r\n"
    "echo 正在启动「表格检索 / 模板改造 / 数据核对」...\r\n"
    "echo 服务地址：http://127.0.0.1:8765/   （关闭本窗口即停止）\r\n"
    "echo.\r\n"
    "%PYEXE% -m app.server --port 8765\r\n"
    "if errorlevel 1 (\r\n"
    "  echo.\r\n"
    "  echo 启动失败，请检查上方错误信息（常见原因：端口被占用 / 缺少 app 目录）。\r\n"
    "  pause\r\n"
    ")\r\n"
)

path = os.path.join(ROOT, "启动.bat")
with open(path, "w", encoding="gbk", newline="") as f:
    f.write(content)
print("written", path, os.path.getsize(path), "bytes")
