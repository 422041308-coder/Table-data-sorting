# 表格检索 / 模板改造 / 数据核对

本地 Web 小工具：对指定文件夹里的所有 Excel 做**全字段模糊检索**、**按模板做字段映射转换**（含"行→列"转置）、**双表差异核对**，并支持**按单元格就地修改**（备份 + 审计 + 回滚）。

- **零第三方依赖**：只用 Python 3.10+ 标准库，离线可用（目标机无网络、无 pandas/openpyxl 也能跑）。
- **体积**：应用代码全部在 `app/` 下，代码量约 1 MB；不需要 Node、数据库、构建工具。
- **数据安全**：默认只读；写回前自动备份、写后自检、失败回滚、全程审计。

> **仓库内 `samples/` 为虚构示例数据**（项目名、单位、人名、电话、合同编号均为占位符），仅用于演示与测试；请勿把真实业务表格提交到仓库。

---

## 一、启动

### 方式 A：EXE 绿色版（推荐，免安装 Python）

- 解压 `dist/表格工具_v1.0_win64.zip` 到任意可写目录（如 `D:\工具\表格工具`），
- 双击 **`表格工具.exe`**：自动启动本地服务并打开浏览器；窗口上可「打开界面 / 打开数据目录 / 停止并退出」。
- 自带精简 Python 运行时，不依赖系统安装 Python；换电脑整目录拷走即可。
- 重新打包：`python scripts/build_exe.py`（详见 [docs/03-EXE打包说明.md](docs/03-EXE打包说明.md)）。

### 方式 B：源码运行（开发/改代码时）

双击 **`启动.bat`**（会自动找 Python，找不到时回退到本机的 DSH Python），浏览器自动打开 <http://127.0.0.1:8765/>。

命令行等价：

```bat
python -m app.server --port 8765
```

首次使用时在顶部把"目标文件夹"切到你放表格的目录（会递归扫描），点「重建索引」。示例目录里运行：索引几十个文件通常几秒到二十秒（只在本进程内），之后每次搜索都是毫秒级。

---

## 二、功能与页面

| 页面 | 能做什么 |
|---|---|
| **检索** | 全字段模糊搜索：多关键词（空格分隔，AND）、短语（英文双引号）、按文件/字段过滤；结果表给出 **文件地址 · 工作表 · 第几行 · 第几列（字母+序号）· A1 · 字段名 · 命中内容高亮**；可打开所在文件夹 / 用 WPS 打开；可导出；点行可直接**编辑并写回**（自动备份+审计） |
| **字段映射** | 登记任意 xlsx 为模板（自动解析多级表头与数据起始格）；选来源表与模板 → 自动推荐字段映射；用**映射矩阵**手工调整规则：`fixed / record / header / span` 选择器 + `cell / block / transpose / filldown / repeat` 模式；**"A 表某行横排 → B 表某列竖排"用 `transpose` 规则**；预演逐单元格确认后保存方案 |
| **模板转换** | 选来源 + 方案 + 模板 → 生成新表：**克隆模板包注入数据，保留原样式/合并/列宽/公式**；输出到 `data/output/` |
| **数据对比** | 左右两张表（可不同文件/工作表）→ 指定主键（1 基列号，可多列组合）→ 输出 新增/删除/修改；**每条差异都带左右两侧的文件/工作表/行/列/A1**；可导出差异报告（修改黄、仅左红、仅右绿） |
| **设置** | 工作区目录、文件清单（格式/大小/状态/错误）、重建索引、修改审计记录 |

---

## 三、已实现的需求对照

| 需求 | 实现 | 位置 |
|---|---|---|
| R1 全字段模糊搜索（所有 Excel 类型） | 魔数识别 + 自研 OOXML / OLE2-BIFF8 / CSV 读取器；行级 bigram 倒排 + 子串校验 | `app/core/sniff.py`、`app/core/readers/`、`app/core/index.py` |
| R2 模板可替换 + 映射行列自由选择（含行→列转置） | 模板登记/替换、坐标化规则引擎、区域转置、模板 zip 克隆注入保样式 | `app/core/template.py`、`app/core/mapping.py`、`app/services/mapping_service.py` |
| R3 双表对比 | 主键匹配（可组合键）+ 字段级差异 + 左右坐标 + 告警 | `app/core/compare.py`、`app/services/compare_service.py` |
| R4 小体积 Web GUI | 原生 HTML/CSS/ES Module 单页，无框架无构建无 CDN | `app/web/`、`app/server.py` |
| R5 结果给出文件地址/第几行第几列 | 全局 `CellRef`（绝对路径/工作表/行列/A1），检索与对比结果都带 | `app/core/model.py` |
| R6 应用内修改 | 最小改动写回（zip 条目级）、冲突检测、备份、自检、审计、撤销 | `app/core/writers/`、`app/core/audit.py`、`app/services/edit_service.py` |

## 四、目录

```
app/
  server.py        HTTP 服务（路由/静态/任务/下载）
  api.py           接口薄壳           mock.py  示例数据（真实服务缺失时回退）
  core/
    model.py       ★数据契约（CellRef / TableModel / MappingRule …）
    sniff.py       魔数嗅探
    readers/       ooxml.py / biff.py / csvr.py（xlsx/xls/et/csv）
    detect.py      多级表头、数据区、小计/占位行识别
    normalize.py   全半角/空白/日期/数值/手机号归一
    index.py       行级倒排索引        search.py  查询解析与高亮
    mapping.py     映射规则引擎        template.py 模板克隆注入
    compare.py     双表差异引擎        xlsxwriter.py 新建 xlsx（导出）
    writers/       xlsx 最小改动写回    audit.py  备份/审计/回滚
  services/        workspace / search / mapping / compare / edit
  web/             index.html app.js style.css page-common.css pages/*.js
data/
  output/          生成的表格           backup/  写回前备份
  audit/           修改审计 jsonl        profiles/ 映射方案
  templates/       模板登记            tmp/     临时
docs/
  CONTRACT.md               契约冻结说明（开发用）
  01-需求整合与项目架构.md     需求 + 架构设计
  02-开发评估与切片方案.md     开发评估与切片
  03-EXE打包说明.md           如何封装成 EXE（已交付绿色版）
  04-操作说明.md              用户手册（含操作说明.html）
  05-软件签名说明.md           代码签名：自签名 / OV / EV / 云签名
  项目架构设计.html           架构书（单文件，浏览器打开）
  操作说明.html               用户手册（单文件，浏览器打开）
  preview/                    界面截图
packaging/
  launcher.cs                EXE 启动器源码（C#，系统自带 csc 编译）
dist/
  表格工具/                   EXE 绿色版（启动器 + 运行时 + 应用）
  表格工具_v1.0_win64.zip     可分发压缩包（13.6 MB）
samples/           样本台账 + 结构导出
scripts/           acceptance.py（验收）、build_exe.py（打包 EXE）、sign_exe.ps1（自签名）、build_html.py（生成 HTML 文档）、probe_xls.py（.xls 探针）
tests/             单元 + 端到端测试（250 项）
启动.bat            源码方式一键启动
```

## 五、测试与验收

```bat
python -m unittest discover -s tests -v        :: 250 项单元/集成测试
python scripts/acceptance.py                  :: 真实语料验收（约 30s，产出 docs/验收报告.md）
```

验收覆盖：全格式读取（66/66 文件）、索引构建、检索坐标、检索导出、同版本零差异、改动版差异带左右坐标、差异报告导出、映射预演、模板转换（保样式 + 行→列转置）、单元格写回 + 审计 + 撤销。当前结果 **13/13 通过**，详见 [验收步骤](#五测试与验收)。

## 六、已知限制

- **.xls / .et(OLE2) 可读不可原地写回**：编辑时提示"另存为 xlsx 后编辑"（老格式写回风险高，见架构文档 D10）。
- 写公式格会清空其旧缓存值，靠 Excel/WPS 打开时重算。
- 撤销按"整文件备份"回滚（批量修改会整批回滚，取首条审计 id）。
- 两级表头若两行之间完全没有纵向合并，可能只认出一行表头（可手动调整数据起始格）。
- 纯数字字符串不按 Excel 日期序列解析（防止编号被误判成日期）。
- 单文件 > 80 MB 或加密文件会被跳过（在「设置」里可见错误原因）。

## 七、契约与扩展

开发接口约定见 [docs/CONTRACT.md](docs/CONTRACT.md)：新增一种表格格式只需在 `app/core/readers/` 加一个适配器并在 `sniff.py` 注册；新增一种映射模式在 `app/core/mapping.py` 的 `_evaluate` 里加分支。`app/core/model.py` 是冻结契约，改动需升 `CONTRACT_VERSION`。
