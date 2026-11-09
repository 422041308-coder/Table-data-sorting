# 封装成 EXE 应用程序：怎么做、做了什么

> 结论：本项目已经打成**免安装绿色版 EXE**，产物在 `dist/`：
> - `dist/表格工具/表格工具.exe` —— 双击即用（同目录下带 `runtime/` 精简 Python 运行时、`app/` 程序）
> - `dist/表格工具_v1.0_win64.zip` —— 13.0 MB 压缩包，解压到任意目录即可用（不依赖系统 Python）
>
> 重新生成：`python scripts/build_exe.py`

---

## 一、封装路线对比

| 路线 | 是否需要联网 | 产物 | 体积 | 适用场景 |
|---|---|---|---|---|
| **A. PyInstaller**（业界标准） | 需要 `pip install pyinstaller` | 单个 exe | 15–25 MB | 有网络、想一个文件搞定 |
| **B. 精简运行时 + 原生启动器**（本项目采用） | **不需要**，全部用系统自带工具 | 文件夹 + 启动 exe（可再 zip） | 36.9 MB / 13 MB 压缩 | 本机无网络、可离线构建、想保留可改代码 |
| **C. 安装包**（Inno Setup / NSIS / IExpress） | Inno/NSIS 需下载；IExpress 系统自带 | 一个安装 exe | 13–15 MB | 要"下一步下一步"式安装、写注册表/开始菜单 |

### 为什么本项目选 B

实测目标机**无网络**：`pip install pyinstaller` 连 pypi/清华/阿里云镜像都失败（TLS 连接被重置），
且系统里没有 PyInstaller/cx_Freeze/Nuitka。但机器上有：

- `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`（系统自带 C# 编译器）
- `C:\Windows\System32\iexpress.exe`（系统自带自解压打包器）
- 本机 DSH 自带的 Python 3.12 运行时（146.9 MB，其中 95.7 MB 是 pandas 等第三方包，与本项目无关）

因此方案 B = **把 Python 运行时裁到最小 + 应用代码 + 一个 C# 写的小启动器**，全程离线、无需下载任何东西。

---

## 二、已交付包的组成

```
dist/表格工具/
├─ 表格工具.exe          10 KB   启动器（C# / WinForms，双击即用）
├─ runtime/              ~35 MB  精简 Python 3.12（python.exe + DLLs + Lib，裁掉 site-packages/test/tkinter 等）
├─ app/                  ~1 MB   应用源码（server/api/core/services/web）
├─ data/                         数据目录（output/backup/audit/profiles/templates/tmp，含一个样本模板登记）
├─ samples/              约 26 KB  样本台账
├─ docs/                 架构书、需求文档、验收报告、界面截图
├─ README.md / 使用说明.txt
└─ 启动.bat              备用启动方式（需要本机有 Python 时）
```

启动器的工作流程（`packaging/launcher.cs`，约 260 行）：

1. 找到同目录 `runtime\python.exe`；
2. 用 `-m app.server --port 8765 -n --log data\logs\service.log` 拉起本地服务（无控制台窗口）；
3. 轮询 `http://127.0.0.1:8765/api/health` 直到就绪（最长 40 秒），然后自动打开默认浏览器；
4. 窗口上提供「打开界面 / 打开数据目录 / 停止并退出」三个按钮；
5. 关闭窗口时结束 python 子进程；若检测到服务已在运行，则直接打开浏览器、不重复启动。

---

## 三、重新打包（一条命令）

```bat
python scripts\build_exe.py
```

脚本做的事（`scripts/build_exe.py`）：

| 步骤 | 内容 |
|---|---|
| 1 | 复制 `app/`（跳过 `__pycache__`） |
| 2 | 建 `data/` 空目录 + 带上已登记的样本模板；**不带**历史备份/审计/输出 |
| 3 | 复制 `samples/`、`docs/`、`README.md`、`启动.bat` |
| 4 | 组装 `runtime/`：复制 python.exe、python312.dll、vcruntime、DLLs/、Lib/（裁掉 `site-packages`、`test`、`ensurepip`、`idlelib`、`tkinter`、`lib2to3`、`pydoc_data`、`venv`、`unittest`、`distutils`、`__pycache__`） |
| 5 | 用系统 `csc.exe` 编译 `packaging/launcher.cs` → `表格工具.exe` |
| 6 | 写 `使用说明.txt`，跑两项自检，zip 打包 |

**自检（构建时自动执行，任一失败即中止）**：

1. `runtime\python.exe -c "import app.server"` —— 精简运行时可用；
2. `表格工具.exe --selftest` —— 启动器能拉起服务、健康检查通过（不弹窗、写完结果文件后退出）。

可换用别的 Python 运行时：

```bat
set EXCEL_TOOLKIT_PYTHON=D:\Python312
python scripts\build_exe.py
```

---

## 四、验证结果（本机实测）

| 项 | 结果 |
|---|---|
| csc 编译启动器 | 通过 |
| 精简运行时自检 | `runtime 3.12.14` |
| 应用导入自检 | `app import ok` |
| 启动器自检（构建时） | `PASS http://127.0.0.1:8765/` |
| **解压 zip 到新目录后独立运行** | 退出码 0，`selftest: PASS http://127.0.0.1:8765/`，`service.log` 里有 `GET /api/health 200` |
| zip 完整性 | 709 个条目、testzip 无错误、包含 exe 与 runtime、不含运行日志 |

> 说明：启动器的**窗口外观**没能在这个开发沙箱里截图验证（沙箱禁止 GUI 程序创建窗口/命名管道，headless Chrome 同样被拦），
> 但启动器除窗口外的全部逻辑（找运行时、拉起服务、健康检查、清理子进程）都用 `--selftest` 验证通过。
> 你本机双击一次即可确认界面。

---

## 五、如果要用 PyInstaller（有网络时）

```bat
pip install pyinstaller

:: 入口脚本 entry.py
::   import os, sys
::   if getattr(sys, "frozen", False):
::       os.environ["EXCEL_TOOLKIT_HOME"] = os.path.dirname(sys.executable)
::       sys.path.insert(0, sys._MEIPASS)
::   from app.server import run
::   run(open_browser=True)

pyinstaller --noconfirm --onefile --noconsole --name 表格工具 ^
  --add-data "app;app" ^
  --collect-submodules app ^
  entry.py
```

- 产物：`dist/表格工具.exe`（单文件，约 15–25 MB）。
- 代码已支持 `EXCEL_TOOLKIT_HOME` 环境变量：设置后 `data/`、备份、审计都落在 exe 旁边，而不是 PyInstaller 的临时解包目录。
- 若把 web 静态资源打进包里，`server.py` 的 `WEB_DIR` 会指向 `_MEIPASS/app/web`，只读可用；写操作都在 `data/` 下，不受影响。

## 六、如果要做"安装包"

- **IExpress（系统自带，离线可用）**：`iexpress /N 安装包.SED`，把 `dist/表格工具` 整个目录作为文件列表，
  `AppLaunched` 指向一个 `install.cmd`（复制到 `%LOCALAPPDATA%\表格工具`、建桌面快捷方式、启动）。
- **Inno Setup / NSIS（推荐，需下载）**：把 `dist/表格工具` 作为 [Files] 源，生成标准 `Setup.exe`，可写开始菜单/卸载项。

## 七、注意事项

1. **体积**：绿色版 36.9 MB 主要来自 Python 运行时（约 35 MB）；想更小就走 PyInstaller 单文件。
2. **杀毒软件**：启动器会拉起同目录的 python.exe，个别安全软件可能拦截；把整个目录加入信任区即可。
3. **端口**：固定 `127.0.0.1:8765`。被占用时启动器会提示；也可命令行指定：
   `runtime\python.exe -m app.server --port 8766`。
4. **数据位置**：所有生成物/备份/审计都在 `data/` 下，换电脑时连同目录一起拷走即可。
5. **只读介质**：若放到 `Program Files` 等无写权限目录，`data/` 会写不进去；建议放 `D:\工具\表格工具` 这类可写目录，或用安装包装到 `%LOCALAPPDATA%`。
6. **老格式**：`.xls/.et` 可读可对比，不支持原地写回（编辑时会提示另存为 xlsx）。
