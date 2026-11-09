# -*- coding: utf-8 -*-
"""生成可上传 GitHub 的脱敏仓库（github-repo/）。

做三件事：
1. 复制源码/文档（排除 dist、data、__pycache__、预览截图、私钥）；
2. 对文本做业务信息脱敏（真实项目名/单位名/人名/电话/合同编码/本机路径/微信目录）；
3. 用同一套映射生成**虚构示例台账** samples/项目台账统计表.xlsx，并重生成结构导出与 HTML 文档。

生成后请在该目录 git init/commit，再用 README 里的步骤 push。
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXPORT = os.path.join(ROOT, "github-repo")

TEXT_EXTS = {".py", ".md", ".js", ".json", ".html", ".txt", ".css", ".cs", ".gitignore"}
BAT_EXTS = {".bat"}
PS1_EXTS = {".ps1"}
SKIP_DIRS = {"dist", "data", "__pycache__", ".git", ".tmp-edge", "github-repo"}
SKIP_FILES = {"codesign.pfx", "codesign.cer", "SHA256SUMS.txt", "install_error.log",
              "install_debug.log", "验收报告.md"}
SKIP_PATHS = {os.path.join("docs", "preview")}

# 顺序很重要：先长后短、先具体后泛化
REPLACEMENTS: list[tuple[str, str]] = [
    # ---------- 本机路径 ----------
    (os.path.join(os.path.expanduser("~"), "Documents"), os.path.join(os.path.expanduser("~"), "Documents")),
    (r"D:\表格", r"D:\表格"),
    (r"C:\Python312",
     r"C:\Python312"),
    (r"C:\Tools", r"C:\Tools"),
    ("示例文件夹", "示例文件夹"),
    ("demo", "demo"),
    ("user", "user"),
    # ---------- 单位/项目（长串优先） ----------
    ("示例项目甲—园区改造工程", "示例项目甲—园区改造工程"),
    ("示例项目甲—园区改造工程", "示例项目甲—园区改造工程"),
    ("示例项目丙—土方清运", "示例项目丙—土方清运"),
    ("2024年示例区示例公司项目台账（单位：元）", "2024年示例区示例公司项目台账（单位：元）"),
    ("2024年示例区示例公司项目台账", "2024年示例区示例公司项目台账"),
    ("示例区", "示例区"),
    ("示例市", "示例市"),
    ("示例", "示例"),
    ("乙", "乙"),
    ("丙", "丙"),
    # ---------- 公司名（含“示例省”，必须早于省名替换） ----------
    ("示例设计院", "示例设计院"),
    ("示例建筑公司", "示例建筑公司"),
    ("示例咨询公司", "示例咨询公司"),
    ("示例审计公司", "示例审计公司"),
    ("示例省", "示例省"),
    # ---------- 合同编码 / 工单号 ----------
    ("DEMO-ZS-24001", "DEMO-ZS-24001"),
    ("DEMO-JS-24002", "DEMO-JS-24002"),
    ("DEMO-JL-24003", "DEMO-JL-24003"),
    ("DEMO-ZX-24004", "DEMO-ZX-24004"),
    ("DEMO-JS-24009", "DEMO-JS-24009"),
    ("DEMO-ZS-24010", "DEMO-ZS-24010"),
    ("DEMO-DX-24002", "DEMO-DX-24002"),
    ("DEMO-YG-24003", "DEMO-YG-24003"),
    ("DEMO", "DEMO"),
    ("demo-zs-24001", "demo-zs-24001"),
    ("demo-js-24002", "demo-js-24002"),
    ("demo-jl-24003", "demo-jl-24003"),
    ("demo-zx-24004", "demo-zx-24004"),
    ("demo-js-24009", "demo-js-24009"),
    ("demo-zs-24010", "demo-zs-24010"),
    ("demo-dx-24002", "demo-dx-24002"),
    ("demo-yg-24003", "demo-yg-24003"),
    ("DEMO-GD-24001", "DEMO-GD-24001"),
    ("DEMO-GD-24002", "DEMO-GD-24002"),
    ("DEMO-GD-24003", "DEMO-GD-24003"),
    ("DEMO-GD-24004", "DEMO-GD-24004"),
    ("DEMO-GD-24005", "DEMO-GD-24005"),
    ("DEMO-GD-240", "DEMO-GD-240"),
    # ---------- 人名 ----------
    ("张三", "张三"), ("李四", "李四"), ("王五", "王五"), ("赵六", "赵六"),
    ("钱七", "钱七"), ("孙八", "孙八"), ("周九", "周九"), ("吴十", "吴十"),
    # ---------- 电话 ----------
    ("13800000001", "13800000001"), ("13800000002", "13800000002"),
    ("13800000003", "13800000003"), ("13800000004", "13800000004"),
    # ---------- 真实语料文件名 ----------
    ("示例台账-修改(1).xlsx", "示例台账-修改(1).xlsx"),
    ("示例台账-修改.xlsx", "示例台账-修改.xlsx"),
    ("示例表格B.xls", "示例表格B.xls"),
    ("示例表格B.xls", "示例表格B.xls"),
    ("示例表格C.xls", "示例表格C.xls"),
    ("示例表格D.xls", "示例表格D.xls"),
    ("示例表格D.xls", "示例表格D.xls"),
    ("示例数据.csv", "示例数据.csv"),
    ("示例台账.xlsx", "示例台账.xlsx"),
    ("示例清单.xlsx", "示例清单.xlsx"),
    ("示例预警.xls", "示例预警.xls"),
    ("示例表格A.xls", "示例表格A.xls"),
    ("示例表格E.et", "示例表格E.et"),
    ("示例表格F.et", "示例表格F.et"),
    ("示例表格G.xlsx", "示例表格G.xlsx"),
    # ---------- 其它 ----------
    ("示例作者", "示例作者"),
]


def should_skip(path: str) -> bool:
    rel = os.path.relpath(path, ROOT)
    parts = rel.split(os.sep)
    if any(p in SKIP_DIRS for p in parts):
        return True
    if any(rel.startswith(sp + os.sep) for sp in SKIP_PATHS):
        return True
    if os.path.basename(path) in SKIP_FILES:
        return True
    if path.endswith((".pyc", ".pyo")):
        return True
    return False


def copy_tree():
    if os.path.isdir(EXPORT):
        for dirpath, _dirnames, filenames in os.walk(EXPORT):
            for name in filenames:
                try:
                    os.chmod(os.path.join(dirpath, name), 0o666)
                except OSError:
                    pass
        shutil.rmtree(EXPORT, ignore_errors=True)
    os.makedirs(EXPORT, exist_ok=True)
    count = 0
    for dirpath, dirnames, filenames in os.walk(ROOT):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for name in filenames:
            src = os.path.join(dirpath, name)
            if should_skip(src):
                continue
            rel = os.path.relpath(src, ROOT)
            dst = os.path.join(EXPORT, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
            try:  # 原样本是只读的，复制后要放开写权限
                os.chmod(dst, 0o666)
            except OSError:
                pass
            count += 1
    print(f"[1/5] 复制文件 {count} 个 → {EXPORT}")


def scrub_text():
    changed = 0
    for dirpath, dirnames, filenames in os.walk(EXPORT):
        for name in filenames:
            path = os.path.join(dirpath, name)
            ext = os.path.splitext(name)[1].lower()
            if ext in BAT_EXTS:
                enc, out_enc = "gbk", "gbk"
                newline = ""
            elif ext in PS1_EXTS:
                enc, out_enc = "utf-8-sig", "utf-8-sig"
                newline = "\r\n"
            elif ext in TEXT_EXTS:
                enc, out_enc = "utf-8-sig", "utf-8"
                newline = None
            else:
                continue
            try:
                raw = open(path, "rb").read()
                text = raw.decode(enc)
            except (UnicodeDecodeError, OSError):
                continue
            orig = text
            for old, new in REPLACEMENTS:
                if old in text:
                    text = text.replace(old, new)
            # 代码里的默认路径改成语义化默认值
            text = text.replace('r"D:\\表格\\示例文件夹"',
                                'os.path.join(os.path.expanduser("~"), "Documents")')
            # 转义形式的本机绝对路径
            text = re.sub(r"C:(?:\\{1,2})Users(?:\\{1,2})user[^\"'\n]*", r"C:\\Python312", text)
            text = text.replace("user", "user")
            # 验收脚本默认扫示例目录
            text = text.replace(
                'CORPUS = os.environ.get("EXCEL_TOOLKIT_CORPUS", os.path.join(ROOT, "samples"))',
                'CORPUS = os.environ.get("EXCEL_TOOLKIT_CORPUS", os.path.join(ROOT, "samples"))')
            # 单测默认没有真实语料 → 置空即自动 skip
            if name == "test_readers.py":
                text = text.replace(
                    'CORPUS = os.environ.get("EXCEL_TOOLKIT_CORPUS", os.path.join(ROOT, "samples"))',
                    'CORPUS = os.environ.get("EXCEL_TOOLKIT_CORPUS", "")')
            if text != orig:
                with open(path, "w", encoding=out_enc, newline=newline) as f:
                    f.write(text)
                changed += 1
    print(f"[2/5] 脱敏文本 {changed} 个文件")


def rebuild_sample():
    """用同一套映射重写示例台账（XML 级替换，保留样式/合并/公式缓存/数组公式）。"""
    import zipfile

    src = os.path.join(ROOT, "samples", "项目台账统计表.xlsx")
    dst = os.path.join(EXPORT, "samples", "项目台账统计表.xlsx")
    if os.path.exists(dst):
        try:
            os.chmod(dst, 0o666)
            os.remove(dst)
        except OSError:
            pass
    hits = 0
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith((".xml", ".rels")):
                text = data.decode("utf-8")
                new = text
                for old, rep in REPLACEMENTS:
                    if old in new:
                        new = new.replace(old, rep)
                if new != text:
                    hits += 1
                    data = new.encode("utf-8")
            zout.writestr(item, data)
    print(f"[3/5] 示例台账已重建（XML 级替换 {hits} 个部件）")


def rebuild_structure():
    """用示例台账重新导出结构说明。"""
    sample = os.path.join(EXPORT, "samples", "项目台账统计表.xlsx")
    out = os.path.join(EXPORT, "samples", "structure.txt")
    script = os.path.join(EXPORT, "scripts", "inspect_sample.py")
    if not (os.path.isfile(script) and os.path.isfile(sample)):
        return
    try:
        subprocess.run([sys.executable, script, sample, out], check=False, timeout=120,
                       cwd=EXPORT, capture_output=True)
        if os.path.isfile(out):
            t = open(out, encoding="utf-8").read()
            t = t.replace(EXPORT, "<repo>").replace(r"C:\Users\user", r"C:\Python312")
            open(out, "w", encoding="utf-8").write(t)
        print("[4/5] structure.txt 已重新生成")
    except Exception as e:  # noqa: BLE001
        print(f"[!] structure.txt 重新生成失败：{e}")


def rebuild_html():
    """重新生成 HTML 文档，并去掉含真实截图的图片引用。"""
    manual_md = os.path.join(EXPORT, "docs", "04-操作说明.md")
    if os.path.isfile(manual_md):
        with open(manual_md, encoding="utf-8") as f:
            lines = f.readlines()
        lines = [ln for ln in lines if not re.search(r"!\[[^\]]*\]\(preview/", ln)]
        with open(manual_md, "w", encoding="utf-8") as f:
            f.writelines(lines)

    gen = os.path.join(EXPORT, "scripts", "build_html.py")
    if not os.path.isfile(gen):
        return
    cmds = [
        [sys.executable, gen,
         os.path.join(EXPORT, "docs", "01-需求整合与项目架构.md"),
         os.path.join(EXPORT, "docs", "02-开发评估与切片方案.md"),
         "--no-kpi", "--title", "表格检索 / 模板改造 / 数据核对 —— 项目架构设计",
         "--subtitle", "脱敏示例版 · 2026-11",
         "-o", os.path.join(EXPORT, "docs", "项目架构设计.html")],
        [sys.executable, gen,
         os.path.join(EXPORT, "docs", "04-操作说明.md"),
         "--no-kpi", "--title", "表格检索 / 模板改造 / 数据核对 —— 操作说明",
         "--subtitle", "用户手册 v1.0 · 脱敏示例版",
         "-o", os.path.join(EXPORT, "docs", "操作说明.html")],
    ]
    for c in cmds:
        try:
            subprocess.run(c, check=False, timeout=120, cwd=EXPORT, capture_output=True)
        except Exception:  # noqa: BLE001
            pass
    print("[5/5] HTML 文档已重新生成（不含真实截图）")


def write_gitignore_readme():
    gitignore = """# 构建产物 / 二进制
dist/
*.exe
*.zip
*.bundle

# 运行时数据（含本地路径、备份、审计、方案）
data/
data/tmp/

# 证书私钥（绝不能提交）
*.pfx
codesign.cer
SHA256SUMS.txt

# Python
__pycache__/
*.pyc
*.pyo
"""
    with open(os.path.join(EXPORT, ".gitignore"), "w", encoding="utf-8") as f:
        f.write(gitignore)

    # 在 README 顶部补一句数据声明
    p = os.path.join(EXPORT, "README.md")
    if os.path.isfile(p):
        with open(p, encoding="utf-8") as f:
            txt = f.read()
        note = (
            "\n> **仓库内 `samples/` 为虚构示例数据**（项目名、单位、人名、电话、合同编号均为占位符），"
            "仅用于演示与测试；请勿把真实业务表格提交到仓库。\n"
        )
        if "虚构示例数据" not in txt:
            txt = txt.replace("\n---\n", note + "\n---\n", 1) if "\n---\n" in txt else note + "\n" + txt
        # 真实语料统计改示例口径
        txt = txt.replace("本机微信接收目录里 66 个表格 / 25 万单元格，索引约 20 秒建好",
                          "示例目录里运行：索引几十个文件通常几秒到二十秒")
        txt = txt.replace("  验收报告.md                 真实语料验收结果\n", "")
        txt = txt.replace("[验收报告](docs/验收报告.md)", "[验收步骤](#五测试与验收)")
        with open(p, "w", encoding="utf-8") as f:
            f.write(txt)
    print("[6/6] .gitignore 与 README 数据声明已写入")


def main():
    print("=" * 60)
    print("  生成脱敏 GitHub 仓库")
    print("=" * 60)
    copy_tree()
    scrub_text()
    rebuild_sample()
    rebuild_structure()
    rebuild_html()
    write_gitignore_readme()
    print("完成：", EXPORT)


if __name__ == "__main__":
    main()
