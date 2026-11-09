# -*- coding: utf-8 -*-
"""打包成 Windows EXE（绿色版）：精简 Python 运行时 + 应用 + C# 启动器。

产物：
    dist/表格工具/表格工具.exe        ← 双击即可运行
    dist/表格工具/runtime/           ← 精简 Python 3.12 运行时（约 45 MB）
    dist/表格工具_v1.0_win64.zip      ← 整个目录的压缩包

说明：
* 本机无网络，PyInstaller 装不上，因此采用"精简运行时 + 原生启动器"方案；
* 启动器用系统自带 csc.exe（.NET Framework 4.0）编译，无需任何下载；
* 运行时不依赖系统安装的 Python，拷到别的 Windows 机器也能直接跑。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

try:  # 控制台可能是 GBK，避免打印特殊字符时崩溃
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:  # noqa: BLE001
    pass

DIST = os.path.join(ROOT, "dist")
STAGE = os.path.join(DIST, "表格工具")
ZIP_PATH = os.path.join(DIST, "表格工具_v1.0_win64.zip")

PY_SRC = os.environ.get(
    "EXCEL_TOOLKIT_PYTHON",
    r"C:\Python312",
)
CSC = r"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

# 运行时裁剪：这些目录/文件对 App 无用，删掉能省 ~70 MB
LIB_EXCLUDE = {
    "site-packages", "test", "ensurepip", "idlelib", "tkinter", "turtledemo",
    "lib2to3", "pydoc_data", "venv", "__pycache__", "unittest", "distutils",
}
COPY_EXCLUDE_DIRS = {"__pycache__"}

README_TXT = """表格检索 / 模板改造 / 数据核对  v1.0
================================================

【怎么用】
1. 双击「表格工具.exe」；
2. 程序会自动启动本地服务并打开浏览器（http://127.0.0.1:8765/）；
3. 顶部把"目标文件夹"切到你放 Excel 的目录，点「重建索引」；
4. 之后就可以搜索、做字段映射转换、双表对比、以及直接改单元格。

【数据在哪】
- 表格文件：你自己指定的目录（程序只读，不会乱动）
- 生成物：本目录下 data\\output\\
- 写回备份：本目录下 data\\backup\\（出问题可从备份还原）
- 修改审计：本目录下 data\\audit\\
- 运行日志：本目录下 data\\logs\\service.log

【注意】
- 关闭「表格工具」窗口即停止服务；想重新用再双击一次。
- 程序自带 Python 运行时，不需要另外安装 Python。
- .xls / .et 老格式可以读取和对比，但修改时会提示"另存为 xlsx"。
- 若杀毒软件拦截，请把本目录加入信任区（启动器会拉起本地 python 进程）。

【杀毒/端口】
- 服务只监听 127.0.0.1:8765，不对外网开放。
- 端口被占用时，关闭占用程序或用命令行指定其它端口：
  runtime\\python.exe -m app.server --port 8766
"""


def log(msg: str):
    print(msg, flush=True)


def clean_dir(path: str):
    if not os.path.isdir(path):
        return
    for attempt in range(3):
        shutil.rmtree(path, ignore_errors=True)
        if not os.path.isdir(path):
            return
        time.sleep(1.0)
    if os.path.isdir(path):
        raise SystemExit(
            f"无法清理旧目录：{path}\n请先关闭正在运行的「表格工具」或占用该目录的程序后重试。"
        )


def copy_app():
    src = os.path.join(ROOT, "app")
    dst = os.path.join(STAGE, "app")
    shutil.copytree(src, dst, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    log(f"[1/6] 已复制应用代码 app/ → {dst}")


def copy_data():
    src = os.path.join(ROOT, "data")
    dst = os.path.join(STAGE, "data")
    os.makedirs(dst, exist_ok=True)
    for name in ("profiles", "templates", "output", "backup", "audit", "tmp"):
        os.makedirs(os.path.join(dst, name), exist_ok=True)
        keep = os.path.join(dst, name, ".gitkeep")
        with open(keep, "w", encoding="utf-8") as f:
            f.write("")
    # 带上已登记的样本模板，开箱即用；其余历史数据不带
    tpl_dir = os.path.join(src, "templates")
    if os.path.isdir(tpl_dir):
        for name in os.listdir(tpl_dir):
            if name.endswith(".json"):
                shutil.copy2(os.path.join(tpl_dir, name), os.path.join(dst, "templates", name))
    log("[2/6] 已准备 data/ 目录（含样本模板、空输出/备份/审计目录）")


def copy_docs():
    for name in ("samples", "docs"):
        src = os.path.join(ROOT, name)
        if os.path.isdir(src):
            shutil.copytree(src, os.path.join(STAGE, name), dirs_exist_ok=True,
                            ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    for name in ("README.md", "启动.bat"):
        src = os.path.join(ROOT, name)
        if os.path.isfile(src):
            shutil.copy2(src, os.path.join(STAGE, name))
    log("[3/6] 已复制样本与文档")


def copy_runtime():
    if not os.path.isdir(PY_SRC):
        raise SystemExit(f"找不到 Python 运行时：{PY_SRC}\n可用环境变量 EXCEL_TOOLKIT_PYTHON 指定。")
    dst = os.path.join(STAGE, "runtime")
    os.makedirs(dst, exist_ok=True)
    # 1) 顶层文件与 DLLs
    for name in os.listdir(PY_SRC):
        full = os.path.join(PY_SRC, name)
        if os.path.isdir(full):
            if name == "DLLs":
                shutil.copytree(full, os.path.join(dst, "DLLs"))
            continue
        if name.lower().endswith((".exe", ".dll", ".txt")):
            shutil.copy2(full, dst)
    # 2) Lib（裁掉不需要的目录）
    src_lib = os.path.join(PY_SRC, "Lib")
    dst_lib = os.path.join(dst, "Lib")
    os.makedirs(dst_lib, exist_ok=True)

    def ignore(dirpath, names):
        skip = set()
        base = os.path.basename(dirpath)
        for n in names:
            if n == "__pycache__":
                skip.add(n)
            elif n in LIB_EXCLUDE and dirpath == src_lib:
                skip.add(n)
            elif n.endswith((".pyc", ".pyo")) and base != "Lib":
                skip.add(n)
        return skip

    for name in os.listdir(src_lib):
        src = os.path.join(src_lib, name)
        if name in LIB_EXCLUDE:
            continue
        dstp = os.path.join(dst_lib, name)
        if os.path.isdir(src):
            shutil.copytree(src, dstp, ignore=ignore)
        elif name.endswith(".py"):
            shutil.copy2(src, dstp)
    log(f"[4/6] 已组装 Python 运行时 runtime/")


def compile_launcher():
    src = os.path.join(ROOT, "packaging", "launcher.cs")
    if not os.path.isfile(src):
        raise SystemExit(f"缺少启动器源码：{src}")
    if not os.path.isfile(CSC):
        raise SystemExit(f"找不到 C# 编译器：{CSC}")
    out = os.path.join(STAGE, "表格工具.exe")
    cmd = [
        CSC, "/nologo", "/target:winexe", "/optimize+", "/codepage:65001",
        f"/out:{out}", src,
        "/reference:System.dll", "/reference:System.Core.dll",
        "/reference:System.Windows.Forms.dll", "/reference:System.Drawing.dll",
    ]
    res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if res.returncode != 0:
        raise SystemExit("C# 编译失败：\n" + (res.stdout or "") + (res.stderr or ""))
    log(f"[5/6] 已编译启动器 表格工具.exe")


def write_notes():
    with open(os.path.join(STAGE, "使用说明.txt"), "w", encoding="utf-8-sig") as f:
        f.write(README_TXT)
    # 启动器图标（可选）：暂无 .ico，跳过
    log("[6/6] 已写入 使用说明.txt")


def verify_runtime():
    py = os.path.join(STAGE, "runtime", "python.exe")
    res = subprocess.run([py, "-c", "import sys; print('runtime', sys.version.split()[0])"],
                         capture_output=True, text=True, encoding="utf-8", errors="replace",
                         cwd=STAGE)
    if res.returncode != 0:
        raise SystemExit("精简运行时自检失败：\n" + (res.stdout or "") + (res.stderr or ""))
    log("      运行时自检：" + (res.stdout or "").strip())
    res2 = subprocess.run([py, "-c", "import app.server; print('app import ok')"],
                          capture_output=True, text=True, encoding="utf-8", errors="replace",
                          cwd=STAGE)
    if res2.returncode != 0:
        raise SystemExit("应用导入失败：\n" + (res2.stdout or "") + (res2.stderr or ""))
    log("      应用导入自检：" + (res2.stdout or "").strip())


def verify_launcher():
    exe = os.path.join(STAGE, "表格工具.exe")
    res_path = os.path.join(STAGE, "data", "logs", "selftest.txt")
    if os.path.isfile(res_path):
        os.remove(res_path)
    log("      启动器自检中（拉起服务并做健康检查，约 5-10 秒）…")
    try:
        subprocess.run([exe, "--selftest"], cwd=STAGE, timeout=90)
    except subprocess.TimeoutExpired:
        raise SystemExit("启动器自检超时（90s）")
    if not os.path.isfile(res_path):
        raise SystemExit("启动器自检没有产出结果文件：" + res_path)
    with open(res_path, encoding="utf-8-sig") as f:
        result = f.read().strip()
    log("      启动器自检：" + result)
    if not result.startswith("PASS"):
        raise SystemExit("启动器自检失败：" + result)


def make_zip():
    if os.path.isfile(ZIP_PATH):
        os.remove(ZIP_PATH)
    t0 = time.time()
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for dirpath, dirnames, filenames in os.walk(STAGE):
            for name in filenames:
                full = os.path.join(dirpath, name)
                rel = os.path.relpath(full, DIST)
                z.write(full, rel)
    log(f"      已生成压缩包：{ZIP_PATH}（{time.time() - t0:.1f}s）")


def dir_size(path: str) -> float:
    total = 0
    for dirpath, _dirs, files in os.walk(path):
        for name in files:
            try:
                total += os.path.getsize(os.path.join(dirpath, name))
            except OSError:
                pass
    return total / (1024 * 1024)


def main():
    t0 = time.time()
    log("=" * 60)
    log("  打包：表格检索 / 模板改造 / 数据核对 → EXE 绿色版")
    log("=" * 60)
    clean_dir(STAGE)
    os.makedirs(STAGE, exist_ok=True)
    copy_app()
    copy_data()
    copy_docs()
    copy_runtime()
    compile_launcher()
    write_notes()
    verify_runtime()
    verify_launcher()
    # 自检产生的日志不随包发布
    logs_dir = os.path.join(STAGE, "data", "logs")
    if os.path.isdir(logs_dir):
        shutil.rmtree(logs_dir, ignore_errors=True)
    make_zip()
    log("")
    log(f"目录：{STAGE}")
    log(f"体积：{dir_size(STAGE):.1f} MB（压缩包 {os.path.getsize(ZIP_PATH) / 1048576:.1f} MB）")
    log(f"耗时：{time.time() - t0:.1f}s")
    log("完成。双击 dist\\表格工具\\表格工具.exe 即可运行。")


if __name__ == "__main__":
    main()
