# 契约冻结说明（CONTRACT v1.0）

> Lead 维护。**所有切片以此为准**；需要改契约先在群里说明，由 Lead 统一改并升 `CONTRACT_VERSION`。
> 数据类定义见 `app/core/model.py`，本文件只约定"谁实现什么、函数长什么样、返回什么"。

---

## 1. 坐标约定（最容易出错，务必遵守）

| 概念 | 基准 | 例子 |
|---|---|---|
| `SheetModel.cells` 的键 | **0 基** `(row, col)` | `(3, 31)` = Excel 的 `AF4` |
| `CellRef.row` / `CellRef.col` | **1 基**，与 Excel 界面一致 | `row=4, col=32` → `AF4` |
| 换算 | 只用 `CellRef.from_index(path, sheet, r0, c0)` | 不要自己 +1/-1 到处写 |
| 列号→字母 | `col_letter(32) == "AF"`；`col_index("AF") == 32` | `app/core/model.py` |
| 合并单元格 | `SheetModel.merges` 里是 0 基闭区间 `(r1, r2, c1, c2)` | `H2:W2` → `(1, 1, 7, 22)` |

API 返回给前端的坐标一律是 `CellRef.to_dict()`：
```json
{"filePath":"C:\\...\\x.xlsx","sheet":"Sheet0","row":12,"col":32,"colLetter":"AF","a1":"AF12","display":"第12行 · 第AF列(第32列) · AF12"}
```

---

## 2. 冻结的文件（只有 Lead 能改）

```
app/core/model.py        数据契约
app/core/xlsxwriter.py   新建 xlsx（导出用），接口 write_xlsx / write_rows_xlsx / auto_widths
app/server.py            HTTP 骨架
app/api.py               API 薄壳（切片不要动，改由 Lead 接）
app/mock.py              示例数据（回退用）
app/web/index.html       前端骨架
app/web/app.js           路由 + API 封装 + el/toast 工具
启动.bat
docs/CONTRACT.md
```

其余文件按切片分工，见第 4 节。**不要碰别人的目录**；测试文件按域分开写。

---

## 3. 模块依赖方向（只能往下依赖）

```
services/*  ──►  core/*  ──►  model.py
                 ▲
                 └── readers/*   (sniff → ooxml / biff / csv)
```

* 不允许 `core` 反向 import `services`；
* 不允许两个切片互相 import（例如 compare 不要 import mapping）；
* 公共工具（归一化、单元格取值）放 `core/normalize.py`，由 S2 维护，其他切片只调用。

---

## 4. 各切片要实现的接口（函数签名冻结）

### S1 采集层（owner: reader）
```python
# app/core/sniff.py
def detect_format(path: str) -> str
    # "xlsx" | "xlsm" | "xlsb" | "xls" | "csv" | "et-zip" | "et-ole" | "unknown"

# app/core/readers/__init__.py
def read_workbook(path: str) -> WorkbookModel   # 内部按 detect_format 分发

# app/core/readers/ooxml.py   def read_ooxml(path: str) -> WorkbookModel
# app/core/readers/biff.py    def read_biff(path: str) -> WorkbookModel
# app/core/readers/csvr.py    def read_csv(path: str) -> WorkbookModel
```
* 单元格：`Cell(raw=..., kind=..., text=..., num=..., formula=...)`；
* 公式：能拿到缓存值就给 `raw/num`，拿不到就给 `text="=..."`、`formula="=..."`、`kind="formula"`；
* 合并区必须填 `SheetModel.merges`；
* 读不了的文件抛 `ValueError("...")`，由上层记 `status=error`，**不要静默返回空表**。

### S2 结构识别与归一化（owner: detect）
```python
# app/core/normalize.py
def normalize_text(s: str) -> str        # 全角→半角、去多余空白、大小写归一（搜索/对比共用）
def cell_display(cell: Cell) -> str       # 展示文本：number 去 .0、date 用 2024.3.9 风格、公式取缓存值
def guess_kind(values: list) -> str       # "text"|"number"|"date"|"phone"|"code"|"money"
def to_number(v) -> float | None
def to_date_text(v) -> str | None

# app/core/detect.py
def detect_table(wb: WorkbookModel, sheet: SheetModel, file_ref: FileRef | None = None) -> TableModel
def detect_workbook(wb: WorkbookModel, file_ref: FileRef | None = None) -> list[TableModel]
```
* 多级表头：字段名用 `"父级 / 子级"` 连接；`header_rows` 用 0 基行号；
* 小计/占位/空行进 `excluded_rows`，不放进 `rows`；`parse_notes` 写判据；
* `row_numbers[i]` 必须与 `rows[i]` 对应（0 基工作表行号），搜索/对比靠它还原坐标。

### S3 检索（owner: search）
```python
# app/core/index.py
class SearchIndex:
    def add(self, table: TableModel) -> None
    def clear(self) -> None
    def size(self) -> dict                       # {"tables": n, "cells": n, "tokens": n}
    def search(self, query: str, limit=50, offset=0, mode="all",
               file_id=None, field=None) -> tuple[list[Hit], int]
    def update_cell(self, ref: CellRef, text: str, field_name="") -> None

# app/core/search.py
def parse_query(q: str) -> dict                  # {"terms": [...], "phrases": [...]}
def highlight(text: str, query: str) -> tuple[str, list[int]]

# app/services/search_service.py
def start_refresh(task, body: dict, root=None) -> dict    # 长任务：更新 task.progress/message
def status(root=None) -> dict
def search(q, file_id=None, field=None, mode="all", limit=50, offset=0, root=None) -> dict
    # → {"total": n, "took": 0.01, "hits": [Hit.to_dict()...], "query": q}
def export(body: dict, root=None, out_dir=None) -> dict   # → {"path": ..., "count": n}
```
* 检索结果必须带完整 `CellRef`；`snippet` 是命中片段，`highlight` 是片段内 `[start, end)`；
* 索引生命周期由 `search_service` 自己管（首次查询时惰性构建）。

### S4 字段映射与模板（owner: mapping）
```python
# app/core/mapping.py
def suggest(source: TableModel, template: TemplateDef) -> list[dict]
    # [{"targetField","sourceField","confidence","reason"}]
def preview(profile: MappingProfile, source: TableModel, template: TemplateDef, limit=50) -> dict
    # {"cells":[MappedCell.to_dict()...], "truncated": bool, "warnings":[...]}
def apply_rules(profile: MappingProfile, source: TableModel, template: TemplateDef) -> dict
    # {(target_row_0based, target_col_0based): value}

# app/core/template.py
def load_template(path: str, sheet: str | None = None) -> TemplateDef
def inject(template_path: str, out_path: str, data_start: tuple[int, int],
           rows: list[dict[int, Any]], sheet: str | None = None) -> str
    # data_start 是 1 基 (row, col)；rows 的键是目标 1 基列号

# app/services/mapping_service.py
def templates() -> list[dict]
def add_template(body: dict, root=None) -> dict
def delete_template(tid: str) -> dict
def set_anchor(tid: str, body: dict) -> dict
def profiles() -> list[dict]
def save_profile(body: dict) -> dict
def delete_profile(pid: str) -> dict
def suggest_by_ids(body: dict, root=None) -> dict      # body: {"fileId","sheetIndex","templateId"}
def preview_by_ids(body: dict, root=None) -> dict      # 同上 + {"limit"}
def start_convert(task, body: dict, root=None) -> dict # 长任务；返回 {"path","count"}
```
* 映射规则见 `MappingRule`：`mode ∈ cell|block|transpose|filldown|repeat`；
  `Sel.kind ∈ fixed|record|header|span`；**行→列转置必须走 `mode="transpose"` 且行列互换**；
* 转换输出必须用 `template.inject`（保样式），不要用 `xlsxwriter` 重画。

### S5 双表对比（owner: compare）
```python
# app/core/compare.py
def compare(left: TableModel, right: TableModel, keys: list[int], options: dict) -> dict
    # → {"summary": {"matched","added","removed","changed","leftRows","rightRows"},
    #    "items": [DiffItem.to_dict()...], "keys": [...], "warnings": [...]}

# app/services/compare_service.py
def compare(body: dict, root=None) -> dict
    # body: {"left":{"path","sheet"},"right":{...},"keys":[32],"options":{"trim":true,"tolerance":0.001,"filldown":true}}
    # keys 是 1 基列号
def export(body: dict, root=None, out_dir=None) -> dict
```
* `DiffItem.kind ∈ added|removed|changed`；added=仅右有，removed=仅左有；
* **每条差异必须带左右 `CellRef`**（有一侧没有时该侧为 null）；
* 主键重复/为空要进 `warnings`，不能静默错配。

### S6 前端（owner: web）
* **只改**：`app/web/pages/*.js`、`app/web/page-*.css`（可选）；
* 页面模块导出：`export async function render(host, ctx)`，`ctx = {api, el, toast, state, app}`；
* 必做页面：`search.js`、`mapping.js`、`convert.js`、`compare.js`、`settings.js`；
* 检索结果列必须含：文件地址（可打开目录/打开文件）、工作表、第几行、第几列（字母+序号）、A1、字段名、命中内容（高亮）、操作（编辑）；
* 对比结果列必须含：类型、主键、字段、左值、左文件/行/列、右值、右文件/行/列；
* 不允许引入任何外部 CDN/框架。

### S8 编辑写回（owner: edit）
```python
# app/core/audit.py
def backup_file(path: str, root: str) -> str            # 返回备份路径
def append_audit(entry: AuditEntry) -> None
def list_audit(file_path: str | None = None, limit=200) -> list[dict]
def undo(op_id: str) -> dict

# app/core/writers/ooxml_writer.py
def write_cell(path: str, ref: CellRef, value, kind="value") -> dict
    # {"ref":..., "oldValue":..., "result":"ok"}；合并区只允许写左上角，否则抛 ValueError
def write_cells(path: str, edits: list) -> int          # 同文件多格：一次备份、一次落盘

# app/services/edit_service.py
def get_cell(file_path, sheet, row, col) -> dict
def write(op: EditOp) -> dict          # 冲突时抛 ApiError("CONFLICT",...,409) 或返回 {"conflict": true}
def write_batch(ops: list[EditOp]) -> dict
def replace_preview(body: dict) -> dict
def replace_apply(body: dict) -> dict
def undo(op_id: str) -> dict
def audit(file=None, limit=200) -> list[dict]
```
* 写回必须：备份 → 最小改动（zip 条目级）→ 重解析自检 → 失败回滚 → 审计；
* 支持的类型：xlsx / xlsm / et-zip；xls / et-ole 抛 `ApiError("UNSUPPORTED_FORMAT", "...")`，由前端提示"另存 xlsx 后编辑"。

---

## 5. 关键 API 响应形状（前端按此渲染）

```jsonc
// GET /api/workspace
{"ok": true, "data": {"root": "C:\\...", "files": 65, "cells": 128000, "indexed": true, "indexedAt": "..."}}

// GET /api/files
{"ok": true, "data": {"root": "...", "files": [{"id","path","name","fmt","size","mtime","sheetsCount","status","error","indexedAt"}]}}

// GET /api/search?q=...
{"ok": true, "data": {"total": 12, "took": 0.004, "query": "...", "hits": [{"ref": {...}, "fieldName": "...", "valueText": "...", "snippet": "...", "highlight": [0, 5], "score": 9.1}]}}

// GET /api/cell
{"ok": true, "data": {"ref": {...}, "value": "...", "kind": "text", "formula": null, "display": "...", "fieldName": "...", "isSubtotal": false, "editable": true, "current": {"mtime": 0, "size": 0}}}

// POST /api/convert  → 长任务
{"ok": true, "data": {"taskId": "abc"}}   // 轮询 GET /api/tasks/abc
// POST /api/compare 可同步返回
```
错误统一：`{"ok": false, "code": "BAD_PARAM", "message": "中文说明", "detail": null}`。

### 5.1 补充约定（v1.1 增补，Lead 仲裁结果）

**文件 id**：`sha1(绝对路径)[:12]`，由 `app/services/workspace_service.py` 的 `file_id(path)` 统一生成。
`/api/files` 返回的 `id` 就是它；搜索/映射/对比接口既可以传 id，也可以直接传绝对路径（服务内部会 `resolve_file`）。

**长任务下载**：`POST /api/convert` 返回 `{taskId}`；轮询 `GET /api/tasks/{id}` 时，若任务已完成且结果含 `path`，
响应会自动补 `result.token` 与 `result.name`，前端用 `/api/download/{token}` 下载。`compare/export`、`search/export` 直接返回 token。

**`/api/open` 白名单**：允许 项目目录、其 `data/`、以及**当前工作区根目录**（用户扫描的那个文件夹）内的路径。

**查找替换**（S8 与前端冻结的 body）：
```jsonc
// POST /api/replace/preview
{"find":"DEMO","replace":"DEMO-NEW","filePath":null,"fileId":null,"matchCase":false,"limit":200}
// → {"ok":true,"data":{"matches":[{"ref":{...},"oldValue":"...","newValue":"..."}],"count":3,"truncated":false}}

// POST /api/replace/apply  —— 只应用用户确认过的条目，逐条带 newValue
{"items":[{"ref":{"filePath":"...","sheet":"...","row":4,"col":32},"newValue":"DEMO-NEW"}]}
// → {"ok":true,"data":{"written":1,"results":[{"ref":{...},"result":"ok"}]}}
```

**映射预演**：`POST /api/mapping/preview` 的 body 在 `{fileId, sheetIndex, templateId, limit}` 之外，
允许附 `rules: [MappingRule.to_dict(), ...]`（未保存方案的前端实时规则），有 `rules` 时优先用 `rules`，否则用 `profileId` 对应方案。


---

## 6. 测试与自测

* 单元测试放 `tests/test_<域>_<主题>.py`，用 `unittest`；
* 运行：`python -m unittest discover -s tests -v`；
* 每个切片交付前必须：自己的单测通过 + 用真实样本跑一遍（样本见 `samples/`，语料见 `D:\表格\示例文件夹`）；
* 不要改别人的测试文件；不要 mock 掉自己核心逻辑。

## 7. 环境约束（务必遵守）

* Windows，Python 3.12，**无网络**：`pip install` 不可用，只能用标准库（`zipfile`/`xml.etree`/`struct`/`sqlite3`/`http.server`/`difflib`/`unicodedata`…）；
* 不要引入 pandas/openpyxl —— 目标机没有，且本项目要求零依赖；
* 源文件默认只读；写回只走 S8；
* 中文输出一律 UTF-8（`encoding="utf-8"`），控制台乱码不代表数据错。
