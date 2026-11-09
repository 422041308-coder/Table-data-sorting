# -*- coding: utf-8 -*-
"""检索：字符二元组倒排 + 子串校验（Lead 维护）。

设计要点：
* 以 **行** 为倒排单元（不是单元格），显著降低 posting 规模；命中后再逐格校验；
* token 统一用"归一化文本的字符 bigram"，中文/英文/数字一视同仁，
  因此任意长度 >= 2 的子串查询都能取到候选；单字符查询退化为全量扫描；
* 最终结果必须做子串校验，避免 bigram 碰撞造成误召回。
"""
from __future__ import annotations

from .model import CellRef, Hit, TableModel

try:  # S2 交付后自动启用；未交付时用本地兜底
    from .normalize import cell_display as _cell_display, normalize_text as _normalize_text
except Exception:  # noqa: BLE001

    def _normalize_text(s: str) -> str:
        import unicodedata

        s = unicodedata.normalize("NFKC", str(s))
        return " ".join(s.split()).lower()

    def _cell_display(cell) -> str:
        if cell is None:
            return ""
        if cell.text:
            return cell.text
        if cell.raw is None:
            return ""
        if isinstance(cell.raw, float) and cell.raw.is_integer():
            return str(int(cell.raw))
        return str(cell.raw)


# 字段权重：命中这些字段名时加权
_FIELD_BOOST = ("项目名称", "合同名称", "合同编码", "中标单位", "供应商", "备注", "标题", "工单编号")


class _RowDoc:
    __slots__ = ("file_path", "sheet", "row0", "cells", "tokens", "joined")

    def __init__(self, file_path: str, sheet: str, row0: int):
        self.file_path = file_path
        self.sheet = sheet
        self.row0 = row0
        self.cells: list[tuple[int, str, str, str]] = []  # (col0, field_name, text, norm)
        self.tokens: set[str] = set()
        self.joined: str = ""  # 该行所有单元格的归一化文本（去掉空格）拼接，用于行级快速判定


def _bigrams(norm: str) -> set[str]:
    """归一化文本 → 字符二元组集合（去空白）。"""
    s = norm.replace(" ", "")
    if len(s) < 2:
        return set()
    return {s[i:i + 2] for i in range(len(s) - 1)}


def _field_boost(name: str) -> float:
    for kw in _FIELD_BOOST:
        if kw in (name or ""):
            return 1.5
    return 0.0


class SearchIndex:
    """行级倒排索引。"""

    def __init__(self):
        self.docs: list[_RowDoc] = []
        self.postings: dict[str, set[int]] = {}
        self.cell_docs: dict[tuple, int] = {}  # (path, sheet, row0, col0) -> doc id
        self.file_rows: dict[str, list[int]] = {}  # path -> doc ids
        self.fields: dict[tuple, str] = {}  # (path,sheet,row0,col0) -> field name
        self.files: list[str] = []
        self._built_at = 0.0

    # ------------------------------------------------------------ 构建
    def clear(self) -> None:
        self.docs.clear()
        self.postings.clear()
        self.cell_docs.clear()
        self.file_rows.clear()
        self.fields.clear()
        self.files.clear()
        self._built_at = 0.0

    def add(self, table: TableModel) -> int:
        """把一个 TableModel 加进索引，返回新增行数。"""
        path = table.path
        if not path:
            return 0
        if path not in self.files:
            self.files.append(path)
        added = 0
        for i, row in enumerate(table.rows):
            row0 = table.row_numbers[i] if i < len(table.row_numbers) else table.data_start + i
            doc = _RowDoc(path, table.sheet, row0)
            for col0, cell in row.items():
                text = _cell_display(cell)
                if not text:
                    continue
                norm = _normalize_text(text)
                if not norm.strip():
                    continue
                field = table.field_at(col0)
                fname = field.name if field else ""
                doc.cells.append((col0, fname, text, norm))
                doc.tokens |= _bigrams(norm)
                key = (path, table.sheet, row0, col0)
                self.fields[key] = fname
                self.cell_docs[key] = len(self.docs)
            if not doc.cells:
                continue
            doc.joined = "\x00".join(c[3].replace(" ", "") for c in doc.cells)
            idx = len(self.docs)
            self.docs.append(doc)
            self.file_rows.setdefault(path, []).append(idx)
            for tok in doc.tokens:
                self.postings.setdefault(tok, set()).add(idx)
            added += 1
        return added

    def size(self) -> dict:
        return {
            "tables": len(self.files),
            "rows": len(self.docs),
            "cells": sum(len(d.cells) for d in self.docs),
            "tokens": len(self.postings),
        }

    # ------------------------------------------------------------ 查询
    def _candidates(self, terms: list[str], mode: str) -> set[int] | None:
        """返回候选行 doc id 集合；None 表示需要全量扫描。"""
        sets: list[set[int]] = []
        for t in terms:
            s = _normalize_text(t).replace(" ", "")
            if len(s) < 2:
                return None
            grams = [s[i:i + 2] for i in range(len(s) - 1)]
            for g in grams:
                posting = self.postings.get(g)
                if not posting:
                    if mode == "any":
                        continue
                    return set()
                sets.append(posting)
        if not sets:
            return None if mode == "any" else set()
        if mode == "any":
            out: set[int] = set()
            for s in sets:
                out |= s
            return out
        # all：从最小集合开始求交
        sets.sort(key=len)
        out = set(sets[0])
        for s in sets[1:]:
            if not out:
                break
            out &= s
        return out

    def search(self, query: str, limit: int = 50, offset: int = 0, mode: str = "all",
               file_id: str | None = None, field: str | None = None,
               path_filter: str | None = None) -> tuple[list[Hit], int]:
        """返回 (当前页命中, 总数)。

        语义：``mode=all`` 表示**同一行内**各词都出现（可分布在不同列）；
        命中的是"包含其中任一词的单元格"，因此同一行可能有多条命中。
        """
        from .search import highlight, parse_query

        parsed = parse_query(query)
        terms = parsed["terms"]
        phrases = parsed["phrases"]
        needles = phrases + terms
        if not needles:
            return [], 0

        cand = self._candidates(needles, "all" if mode != "any" else "any")
        scan_ids = range(len(self.docs)) if cand is None else sorted(cand)

        # 预先归一化查询词，避免在行/单元格循环里反复做 NFKC 与空格处理
        norm_needles = [(_normalize_text(n).replace(" ", ""), n) for n in needles]

        hits: list[Hit] = []
        for doc_id in scan_ids:
            doc = self.docs[doc_id]
            if path_filter and doc.file_path != path_filter:
                continue
            # 字段过滤后，行级判定只看该字段内的单元格
            view = [c for c in doc.cells if not field or field in c[1]]
            if not view:
                continue
            joined = doc.joined if not field else "\x00".join(
                c[3].replace(" ", "") for c in view)
            matched = [orig for qn, orig in norm_needles if qn in joined]
            if mode == "any":
                if not matched:
                    continue
            elif mode == "phrase":
                if len(matched) < len(phrases) and phrases:
                    continue
                if not matched:
                    continue
            else:  # all
                if len(matched) < len(needles):
                    continue
            for col0, fname, text, norm in view:
                cell_norm = norm.replace(" ", "")
                hit_needles = [orig for qn, orig in norm_needles if qn in cell_norm]
                if not hit_needles:
                    continue
                snippet, marks = highlight(text, hit_needles[0], terms)
                score = self._score(text, fname, hit_needles, norm_needles)
                ref = CellRef.from_index(doc.file_path, doc.sheet, doc.row0, col0)
                hits.append(Hit(ref=ref, field_name=fname, value_text=text,
                                snippet=snippet, highlight=marks, score=score))
        hits.sort(key=lambda h: (-h.score, h.ref.file_path, h.ref.sheet, h.ref.row, h.ref.col))
        total = len(hits)
        return hits[offset:offset + limit], total

    @staticmethod
    def _score(text: str, fname: str, needles: list[str], norm_needles: list[tuple[str, str]]) -> float:
        score = 1.0
        cell_norm_cache: str | None = None
        for qn, _orig in norm_needles:
            if _orig not in needles:
                continue
            if cell_norm_cache is None:
                cell_norm_cache = _normalize_text(text).replace(" ", "")
            pos = cell_norm_cache.find(qn)
            if pos == 0:
                score += 3.0
            elif pos > 0:
                score += 1.0
            if cell_norm_cache == qn:
                score += 2.0
        score += _field_boost(fname)
        score += max(0.0, 2.0 - len(text) / 80.0)
        return round(score, 3)

    # ------------------------------------------------------------ 增量
    def update_cell(self, ref: CellRef, text: str, field_name: str = "") -> bool:
        """写回后增量更新一格。找不到返回 False。"""
        key = (ref.file_path, ref.sheet, ref.row - 1, ref.col - 1)
        doc_id = self.cell_docs.get(key)
        if doc_id is None:
            return False
        doc = self.docs[doc_id]
        for i, (col0, fname, old_text, old_norm) in enumerate(doc.cells):
            if col0 == ref.col - 1:
                norm = _normalize_text(text)
                doc.cells[i] = (col0, field_name or fname, text, norm)
                break
        else:
            return False
        # 重建该行的 token 与 posting
        for tok in doc.tokens:
            posting = self.postings.get(tok)
            if posting is not None:
                posting.discard(doc_id)
                if not posting:
                    self.postings.pop(tok, None)
        doc.tokens = set()
        for _c, _f, t, _n in doc.cells:
            doc.tokens |= _bigrams(_n)
        doc.joined = "\x00".join(c[3].replace(" ", "") for c in doc.cells)
        for tok in doc.tokens:
            self.postings.setdefault(tok, set()).add(doc_id)
        return True

    # ------------------------------------------------------------ 清理
    def remove_file(self, path: str) -> None:
        ids = set(self.file_rows.get(path, []))
        if not ids:
            return
        for doc_id in ids:
            doc = self.docs[doc_id]
            for tok in doc.tokens:
                posting = self.postings.get(tok)
                if posting is not None:
                    posting.discard(doc_id)
                    if not posting:
                        self.postings.pop(tok, None)
            for col0, _fn, _t, _n in doc.cells:
                self.cell_docs.pop((path, doc.sheet, doc.row0, col0), None)
                self.fields.pop((path, doc.sheet, doc.row0, col0), None)
        # 索引重排：doc id 会失效，简单起见直接整体重建由上层负责
        self.file_rows.pop(path, None)
        if path in self.files:
            self.files.remove(path)
