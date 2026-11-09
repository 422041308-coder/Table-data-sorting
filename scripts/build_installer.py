# -*- coding: utf-8 -*-
"""生成"自安装包"（单个 EXE，内置程序 + Python 运行时 + 证书 + 卸载器）。

原理：Windows 自带 IExpress（自解压打包器）+ install.cmd，全程离线。
产物：
    dist/表格工具_安装包_v1.0.exe    单个安装程序（约 13 MB）
安装内容：
    %LOCALAPPDATA%\\表格工具\\        程序 + runtime（Python 运行时）+ codesign.cer + 卸载入口
    桌面/开始菜单快捷方式、HKCU 卸载项、当前用户受信任根证书（自签名信任）

用法：
    python scripts/build_installer.py
    python scripts/build_installer.py --install-dir "D:\\工具\\表格工具"
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIST = os.path.join(ROOT, "dist")
APP_DIR = os.path.join(DIST, "表格工具")
STAGE = os.path.join(DIST, "installer")
OUT_EXE = os.path.join(DIST, "表格工具_安装包_v1.0.exe")
IEXPRESS = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "iexpress.exe")
CERT = os.path.join(DIST, "codesign.cer")

DEFAULT_INSTALL_DIR = r"%LOCALAPPDATA%\表格工具"

try:
    sys.stdout.reconfigure(errors="replace")
except Exception:  # noqa: BLE001
    pass


def log(msg: str):
    print(msg, flush=True)


def gbk(path: str, content: str):
    """写 GBK + CRLF（cmd 批处理需要）。"""
    with open(path, "w", encoding="gbk", newline="") as f:
        f.write(content.replace("\n", "\r\n"))


def write_install_cmd():
    content = r"""@echo off
setlocal EnableExtensions
title 表格工具 安装程序

set "SRC=%~dp0"
set "TARGET=__INSTALL_DIR__"
if not "%INSTALL_DIR%"=="" set "TARGET=%INSTALL_DIR%"
set "SILENT="
if /I "%~1"=="/S" set "SILENT=1"
if /I "%~1"=="/silent" set "SILENT=1"

echo ============================================
echo   表格检索 / 模板改造 / 数据核对  安装程序
echo ============================================
echo 安装位置: %TARGET%
echo.

if not exist "%TARGET%" mkdir "%TARGET%" 2>nul

echo [1/5] 解压程序文件（含 Python 运行时）...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -LiteralPath '%SRC%payload.zip' -DestinationPath '%TARGET%' -Force" >nul 2>nul
if not exist "%TARGET%\runtime\python.exe" (
  echo   解压失败：请确认以普通用户身份运行、且目标目录可写。
  if not defined SILENT pause
  exit /b 1
)
echo   完成

copy /y "%SRC%uninstall.ps1" "%TARGET%\uninstall.ps1" >nul 2>nul
copy /y "%SRC%uninstall.bat" "%TARGET%\uninstall.bat" >nul 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "Copy-Item -LiteralPath '%TARGET%\uninstall.bat' -Destination '%TARGET%\卸载.bat' -Force" >nul 2>nul

echo [2/5] 安装自签名证书（当前用户受信任根 + 受信任发布者）...
if exist "%TARGET%\codesign.cer" (
  certutil -user -addstore -f Root "%TARGET%\codesign.cer" >nul 2>nul
  certutil -user -addstore -f TrustedPublisher "%TARGET%\codesign.cer" >nul 2>nul
  echo   完成（若安装期间弹出安全警告，请点"是"）
) else (
  echo   未找到证书，跳过
)

echo [3/5] 创建桌面与开始菜单快捷方式...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $desk=[Environment]::GetFolderPath('Desktop'); $s=$ws.CreateShortcut((Join-Path $desk '表格工具.lnk')); $s.TargetPath='%TARGET%\表格工具.exe'; $s.WorkingDirectory='%TARGET%'; $s.Description='表格检索/模板改造/数据核对'; $s.Save(); $menu=[Environment]::GetFolderPath('Programs'); $s2=$ws.CreateShortcut((Join-Path $menu '表格工具.lnk')); $s2.TargetPath='%TARGET%\表格工具.exe'; $s2.WorkingDirectory='%TARGET%'; $s2.Save()" >nul 2>nul
echo   完成

echo [4/5] 写入卸载信息...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\表格工具" /v "DisplayName" /d "表格工具（检索/模板/核对）" /f >nul 2>nul
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\表格工具" /v "DisplayVersion" /d "1.0" /f >nul 2>nul
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\表格工具" /v "Publisher" /d "杨子江 by DeepSeek HARNESS" /f >nul 2>nul
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\表格工具" /v "InstallLocation" /d "%TARGET%" /f >nul 2>nul
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\表格工具" /v "UninstallString" /d "\"%TARGET%\uninstall.bat\"" /f >nul 2>nul
echo   完成

echo [5/5] 启动程序...
if defined SILENT (
  echo   静默模式：已跳过自动启动
) else (
  start "" "%TARGET%\表格工具.exe"
)

echo.
echo 安装完成。以后从桌面「表格工具」图标启动；卸载请运行安装目录下的 卸载.bat。
if not defined SILENT (
  echo.
  pause
)
exit /b 0
"""
    content = content.replace("__INSTALL_DIR__", DEFAULT_INSTALL_DIR)
    gbk(os.path.join(STAGE, "install.cmd"), content)


UNINSTALL_PS1 = r"""# 表格工具 卸载脚本（由安装包生成）
$ErrorActionPreference = "SilentlyContinue"
$Target = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "正在卸载：表格工具"
Write-Host "安装目录：$Target"

# 1) 关闭正在运行的进程
Get-Process -Name "表格工具" | Stop-Process -Force

# 2) 删除证书（按主体匹配）
foreach ($store in @("Root", "TrustedPublisher")) {
    Get-ChildItem "Cert:\CurrentUser\$store" |
        Where-Object { $_.Subject -like "*杨子江*" } |
        ForEach-Object {
            certutil -user -delstore $store $_.Thumbprint | Out-Null
            Write-Host "已删除证书：$store $($_.Thumbprint)"
        }
}

# 3) 删除快捷方式
foreach ($p in @(
    (Join-Path ([Environment]::GetFolderPath("Desktop")) "表格工具.lnk"),
    (Join-Path ([Environment]::GetFolderPath("Programs")) "表格工具.lnk")
)) { if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force } }

# 4) 删除注册表卸载项
Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\表格工具" -Recurse -Force

# 5) 延时删除安装目录（脚本自身在目录里，先复制到临时目录执行）
Write-Host "正在删除程序文件…"
$tmp = Join-Path $env:TEMP ("uninstall_表格工具_" + [guid]::NewGuid().ToString("N") + ".cmd")
$cmd = '@echo off' + "`r`n" + 'timeout /t 2 >nul' + "`r`n" + 'rmdir /s /q "' + $Target + '"' + "`r`n" + 'del /f /q "%~f0" >nul 2>nul'
[System.IO.File]::WriteAllText($tmp, $cmd, [System.Text.Encoding]::Default)
Start-Process -FilePath $tmp -WindowStyle Hidden
Write-Host "卸载完成（程序目录会随后自动删除）。"
Start-Sleep -Seconds 3
"""

UNINSTALL_BAT = r"""@echo off
chcp 65001 >nul
title 卸载 表格工具
echo 即将卸载「表格工具」（程序目录、快捷方式、自签名证书）。
choice /c YN /m "确认卸载"
if errorlevel 2 exit /b 0
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
pause
"""


def make_payload():
    if not os.path.isdir(APP_DIR):
        raise SystemExit(f"找不到 {APP_DIR}，请先运行 scripts/build_exe.py")
    payload = os.path.join(STAGE, "payload.zip")
    if os.path.isfile(payload):
        os.remove(payload)
    with zipfile.ZipFile(payload, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for dirpath, _dirs, files in os.walk(APP_DIR):
            for name in files:
                full = os.path.join(dirpath, name)
                rel = os.path.relpath(full, APP_DIR)
                z.write(full, rel)
        # 证书随包安装
        if os.path.isfile(CERT):
            z.write(CERT, "codesign.cer")
    log(f"  已生成 payload.zip（{os.path.getsize(payload) / 1048576:.1f} MB）")
    return payload


def write_sed(sed_path: str, payload: str):
    sed = r"""[Version]
Class=IEXPRESS
SEDVersion=3

[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=0
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=%InstallPrompt%
DisplayLicense=%DisplayLicense%
FinishMessage=%FinishMessage%
TargetName=%TargetName%
FriendlyName=%FriendlyName%
AppLaunched=%AppLaunched%
PostInstallCmd=%PostInstallCmd%
AdminQuietInstCmd=%AdminQuietInstCmd%
UserQuietInstCmd=%UserQuietInstCmd%
SourceFiles=SourceFiles

[Strings]
InstallPrompt=
DisplayLicense=
FinishMessage=
TargetName=__TARGET__
FriendlyName=表格工具 安装程序
AppLaunched=install.cmd
PostInstallCmd=<None>
AdminQuietInstCmd=
UserQuietInstCmd=
FILE0="payload.zip"
FILE1="install.cmd"

[SourceFiles]
SourceFiles0=__STAGE__
[SourceFiles0]
%FILE0%=
%FILE1%=
"""
    sed = sed.replace("__TARGET__", OUT_EXE).replace("__STAGE__", STAGE + os.sep)
    # IExpress 读 ANSI/GBK，中文必须用 GBK 保存
    with open(sed_path, "w", encoding="gbk", newline="\r\n") as f:
        f.write(sed)


def build():
    t0 = time.time()
    log("=" * 60)
    log("  生成自安装包（内置程序 + Python 运行时 + 证书 + 卸载器）")
    log("=" * 60)
    if os.path.isdir(STAGE):
        shutil.rmtree(STAGE, ignore_errors=True)
    os.makedirs(STAGE, exist_ok=True)

    log("[1/4] 打包 payload…")
    make_payload()

    log("[2/4] 生成安装/卸载脚本…")
    write_install_cmd()
    gbk(os.path.join(STAGE, "uninstall.ps1"), UNINSTALL_PS1.replace("__INSTALL_DIR__", DEFAULT_INSTALL_DIR))
    gbk(os.path.join(STAGE, "卸载.bat"), UNINSTALL_BAT)
    gbk(os.path.join(STAGE, "uninstall.bat"), UNINSTALL_BAT)

    log("[3/4] 用 csc 编译自解压安装器…")
    csc = r"C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\csc.exe"
    src = os.path.join(ROOT, "packaging", "installer.cs")
    cmd = [csc, "/nologo", "/target:winexe", "/optimize+", "/codepage:65001",
           f"/out:{OUT_EXE}", src,
           "/reference:System.dll", "/reference:System.Windows.Forms.dll",
           "/reference:System.Drawing.dll",
           f"/resource:{os.path.join(STAGE, 'payload.zip')}",
           f"/resource:{os.path.join(STAGE, 'install.cmd')}",
           f"/resource:{os.path.join(STAGE, 'uninstall.ps1')}",
           f"/resource:{os.path.join(STAGE, '卸载.bat')}",
           f"/resource:{os.path.join(STAGE, 'uninstall.bat')}"]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if not os.path.isfile(OUT_EXE):
        raise SystemExit("安装器编译失败：\\n" + (r.stdout or "") + (r.stderr or ""))
    log(f"  已生成：{OUT_EXE}（{os.path.getsize(OUT_EXE) / 1048576:.1f} MB）")

    log("[4/4] 生成 SHA256 校验和…")
    sums = os.path.join(DIST, "SHA256SUMS.txt")
    lines = []
    for name in ("表格工具.exe", "表格工具_v1.0_win64.zip"):
        p = os.path.join(DIST, "表格工具", name) if name == "表格工具.exe" else os.path.join(DIST, name)
        if os.path.isfile(p):
            import hashlib

            h = hashlib.sha256(open(p, "rb").read()).hexdigest()
            lines.append(f"{h}  {name}")
    import hashlib

    lines.append(f"{hashlib.sha256(open(OUT_EXE, 'rb').read()).hexdigest()}  {os.path.basename(OUT_EXE)}")
    with open(sums, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    log("")
    log(f"完成，耗时 {time.time() - t0:.1f}s")
    log(f"安装包：{OUT_EXE}")
    log("测试：在命令行运行 \"表格工具_安装包_v1.0.exe\" /S 可静默安装到 %LOCALAPPDATA%\\表格工具")


if __name__ == "__main__":
    build()
