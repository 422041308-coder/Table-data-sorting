# -*- coding: utf-8 -*-
"""查询解析与命中片段高亮（Lead 维护）。"""
from __future__ import annotations

_QUOTE_RE = None


def parse_query(q: str) -> dict:
    """把用户输入拆成短语（引号内）与普通词。

    支持 ``"..."``、``“...”``、``「...」`` 三种引号。
    """
    global _QUOTE_RE
    import re

    if _QUOTE_RE is None:
        _QUOTE_RE = re.compile(r'["“「]([^"”」]+)["”」]')
    q = (q or "").strip()
    phrases = [m.group(1).strip() for m in _QUOTE_RE.finditer(q) if m.group(1).strip()]
    rest = _QUOTE_RE.sub(" ", q)
    terms = [t for t in rest.split() if t]
    return {"terms": terms, "phrases": phrases, "raw": q}


def highlight(text: str, needle: str, terms: list[str] | None = None,
              context: int = 24) -> tuple[str, list[int]]:
    """返回 (命中片段, [start, end])；start/end 是片段内的相对位置。"""
    if not text:
        return "", []
    idx = text.find(needle)
    used = needle
    if idx < 0 and needle:
        idx = text.lower().find(needle.lower())
    if idx < 0:
        for t in (terms or []):
            if not t:
                continue
            idx = text.find(t)
            used = t
            if idx < 0:
                idx = text.lower().find(t.lower())
            if idx >= 0:
                break
    if idx < 0:
        head = text[:60]
        return (head + "…") if len(text) > 60 else head, []
    start = max(0, idx - context)
    end = min(len(text), idx + len(used) + context)
    prefix = "…" if start > 0 else ""
    suffix = "…" if end < len(text) else ""
    snippet = prefix + text[start:end] + suffix
    hs = idx - start + len(prefix)
    return snippet, [hs, hs + len(used)]
