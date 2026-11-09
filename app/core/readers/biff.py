# -*- coding: utf-8 -*-
"""BIFF8（.xls / WPS .et OLE 版）读取器：纯标准库实现 CFB + BIFF8。

覆盖：
* CFB（OLE2 复合文档）：FAT / DIFAT / MiniFAT / 目录 / 流读取；
* 全局子流：BOF 版本、BOUNDSHEET（表名/隐藏）、SST（含 CONTINUE 分片）、
  FORMAT + XF（日期格式判定）、1904 日期体系；
* 工作表子流：DIMENSIONS、LABELSST、LABEL、RSTRING、NUMBER、RK、MULRK、
  BOOLERR、FORMULA（缓存值 + 后续 STRING 记录）、MERGEDCELLS、BLANK/MULBLANK。

坐标：cells 键为 0 基 ``(row, col)``；merges 为 0 基闭区间 ``(r1, r2, c1, c2)``。
解析失败一律抛 ``ValueError``（中文原因），不静默返回空表。
"""
from __future__ import annotations

import os
import struct

from ..model import (
    KIND_BOOL,
    KIND_DATE,
    KIND_FORMULA,
    KIND_NUMBER,
    KIND_TEXT,
    Cell,
    SheetModel,
    WorkbookModel,
)
from .ooxml import classify_date_format, format_serial_date, _format_number

_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_MAXREGSECT = 0xFFFFFFFA
_ENDOFCHAIN = 0xFFFFFFFE

# 单元格记录编号
_RID_BOF = 0x0809
_RID_EOF = 0x000A
_RID_DIMENSIONS = 0x0200
_RID_BLANK = 0x0201
_RID_NUMBER = 0x0203
_RID_LABEL = 0x0204
_RID_BOOLERR = 0x0205
_RID_STRING = 0x0207
_RID_ROW = 0x0208
_RID_RK = 0x027E
_RID_MULRK = 0x00BD
_RID_MULBLANK = 0x00BE
_RID_LABELSST = 0x00FD
_RID_FORMULA = 0x0006
_RID_MERGEDCELLS = 0x00E5
_RID_RSTRING = 0x00D6
_RID_CONTINUE = 0x003C
# 全局子流记录编号
_RID_BUNDLESHEET = 0x0085
_RID_SST = 0x00FC
_RID_XF = 0x00E0
_RID_FORMAT = 0x041E
_RID_1904 = 0x0022
_RID_FILEPASS = 0x002F

_ERROR_TEXT = {
    0x00: "#NULL!",
    0x07: "#DIV/0!",
    0x0F: "#VALUE!",
    0x17: "#REF!",
    0x1D: "#NAME?",
    0x24: "#NUM!",
    0x2A: "#N/A",
}


# ==========================================================================
# 二进制小工具
# ==========================================================================
def _u16(buf: bytes, off: int = 0) -> int:
    if off + 2 > len(buf):
        raise ValueError("数据越界（期望 2 字节）")
    return struct.unpack_from("<H", buf, off)[0]


def _u32(buf: bytes, off: int = 0) -> int:
    if off + 4 > len(buf):
        raise ValueError("数据越界（期望 4 字节）")
    return struct.unpack_from("<I", buf, off)[0]


def _u64(buf: bytes, off: int = 0) -> int:
    if off + 8 > len(buf):
        raise ValueError("数据越界（期望 8 字节）")
    return struct.unpack_from("<Q", buf, off)[0]


# ==========================================================================
# CFB（OLE2 复合文档）
# ==========================================================================
class CFB:
    """最小可用的 OLE2 / Compound File Binary 读取器（BIFF 工作簿流）。

    ``entries`` 暴露目录项：``{"name", "type", "start", "size"}``，type=2 为流。
    """

    def __init__(self, data: bytes, name: str = ""):
        self.data = data
        self.name = name or "OLE2 文件"
        if len(data) < 512 or not data.startswith(_OLE_MAGIC):
            raise ValueError(f"{self.name} 不是 OLE2 复合文档")
        self.ssz = 1 << _u16(data, 30)
        self.mssz = 1 << _u16(data, 32)
        if self.ssz not in (512, 4096) or self.mssz != 64:
            raise ValueError(f"{self.name} 的 OLE2 扇区大小异常"
                             f"（sector={self.ssz}, mini={self.mssz}）")
        self.n_fat = _u32(data, 44)
        self.first_dir = _u32(data, 48)
        self.mini_cutoff = _u32(data, 56)
        self.first_minifat = _u32(data, 60)
        self.n_minifat = _u32(data, 64)
        self.first_difat = _u32(data, 68)
        self.n_difat = _u32(data, 72)
        self.n_sectors = max(0, (len(data) - self.ssz) // self.ssz)
        self.fat = self._load_fat()
        self.entries = self._load_directory()
        self.root = next((e for e in self.entries if e["type"] == 5), None)
        if self.root is None:
            raise ValueError(f"{self.name} 的 OLE2 目录缺少根条目")
        self.minifat = self._load_minifat()
        self.ministream = b""
        if self.root["size"] and self.root["start"] < _MAXREGSECT:
            self.ministream = self._read_fat_chain(self.root["start"], self.root["size"])

    # -- 扇区与链 ----------------------------------------------------------
    def _sector(self, sid: int) -> bytes:
        if sid > _MAXREGSECT:
            raise ValueError(f"{self.name} 的 OLE2 扇区号非法：0x{sid:08X}")
        off = (sid + 1) * self.ssz
        if off >= len(self.data):
            raise ValueError(f"{self.name} 的 OLE2 数据被截断（扇区 {sid} 越界）")
        return self.data[off:off + self.ssz]

    def _load_fat(self) -> list[int]:
        difat: list[int] = []
        for i in range(109):
            sid = _u32(self.data, 76 + i * 4)
            if sid >= _MAXREGSECT:
                break
            difat.append(sid)
        sid = self.first_difat
        seen: set[int] = set()
        while sid < _MAXREGSECT and len(difat) < self.n_fat:
            if sid in seen:
                raise ValueError(f"{self.name} 的 OLE2 DIFAT 链存在环")
            seen.add(sid)
            sector = self._sector(sid)
            count = self.ssz // 4 - 1
            for k in range(count):
                extra = _u32(sector, k * 4)
                if extra < _MAXREGSECT:
                    difat.append(extra)
            sid = _u32(sector, count * 4)
        fat: list[int] = []
        for fat_sid in difat[:self.n_fat]:
            sector = self._sector(fat_sid)
            if len(sector) < self.ssz:
                raise ValueError(f"{self.name} 的 OLE2 FAT 扇区被截断")
            count = self.ssz // 4
            fat.extend(struct.unpack("<%dI" % count, sector))
        if not fat:
            raise ValueError(f"{self.name} 的 OLE2 FAT 为空（文件损坏）")
        return fat

    def _load_directory(self) -> list[dict]:
        entries: list[dict] = []
        sid = self.first_dir
        seen: set[int] = set()
        while sid < _MAXREGSECT and sid not in seen:
            seen.add(sid)
            sector = self._sector(sid)
            for k in range(len(sector) // 128):
                raw = sector[k * 128:(k + 1) * 128]
                if len(raw) < 128:
                    break
                name_len = _u16(raw, 64)
                entry_type = raw[66]
                if entry_type == 0 and name_len == 0:
                    continue
                name = ""
                if name_len >= 2:
                    name = raw[:name_len - 2].decode("utf-16-le", "ignore")
                entries.append({
                    "name": name,
                    "type": entry_type,
                    "start": _u32(raw, 116),
                    "size": _u64(raw, 120),
                })
            sid = self.fat[sid] if sid < len(self.fat) else _ENDOFCHAIN
        return entries

    def _load_minifat(self) -> list[int]:
        if not self.n_minifat or self.first_minifat >= _MAXREGSECT:
            return []
        raw = self._read_fat_chain(self.first_minifat, self.n_minifat * self.ssz)
        count = len(raw) // 4
        return list(struct.unpack("<%dI" % count, raw[:count * 4]))

    def _read_fat_chain(self, start: int, size: int) -> bytes:
        if size <= 0:
            return b""
        parts: list[bytes] = []
        total = 0
        sid = start
        seen: set[int] = set()
        while total < size:
            if sid >= _MAXREGSECT:
                break
            if sid in seen:
                raise ValueError(f"{self.name} 的 OLE2 数据链存在环")
            seen.add(sid)
            if sid >= len(self.fat):
                raise ValueError(f"{self.name} 的 OLE2 FAT 索引越界：{sid}")
            sector = self._sector(sid)
            parts.append(sector)
            total += len(sector)
            sid = self.fat[sid]
        return b"".join(parts)[:size]

    def _read_mini_chain(self, start: int, size: int) -> bytes:
        if size <= 0:
            return b""
        parts: list[bytes] = []
        total = 0
        sid = start
        seen: set[int] = set()
        while total < size:
            if sid >= _MAXREGSECT:
                break
            if sid in seen:
                raise ValueError(f"{self.name} 的 OLE2 迷你链存在环")
            seen.add(sid)
            off = sid * self.mssz
            if off >= len(self.ministream):
                raise ValueError(f"{self.name} 的 OLE2 迷你流越界：{sid}")
            parts.append(self.ministream[off:off + self.mssz])
            total += self.mssz
            sid = self.minifat[sid] if sid < len(self.minifat) else _ENDOFCHAIN
        return b"".join(parts)[:size]

    # -- 对外 --------------------------------------------------------------
    def stream_names(self) -> list[str]:
        return [e["name"] for e in self.entries if e["type"] == 2]

    def read_stream(self, name: str) -> bytes:
        for entry in self.entries:
            if entry["type"] == 2 and entry["name"].lower() == name.lower():
                size = entry["size"]
                if size == 0:
                    return b""
                if size < self.mini_cutoff:
                    return self._read_mini_chain(entry["start"], size)
                return self._read_fat_chain(entry["start"], size)
        raise KeyError(name)


# ==========================================================================
# BIFF 基础解析
# ==========================================================================
def _iter_records(data: bytes, start: int = 0):
    """按 BIFF 记录结构 (id, body) 迭代。"""
    pos = start
    size = len(data)
    while pos + 4 <= size:
        rid, length = struct.unpack_from("<HH", data, pos)
        body = data[pos + 4:pos + 4 + length]
        yield rid, body
        pos += 4 + length


class _ChunkReader:
    """SST 分片读取：跨 CONTINUE 记录时处理字符编码标志字节。"""

    __slots__ = ("chunks", "ci", "pos")

    def __init__(self, chunks: list[bytes]):
        self.chunks = chunks
        self.ci = 0
        self.pos = 0

    def _advance(self) -> None:
        while self.ci < len(self.chunks) and self.pos >= len(self.chunks[self.ci]):
            self.ci += 1
            self.pos = 0

    def _at_end(self) -> bool:
        self._advance()
        return self.ci >= len(self.chunks)

    def read(self, count: int) -> bytes:
        out = bytearray()
        while count > 0:
            if self._at_end():
                raise ValueError("共享字符串表数据意外结束（CONTINUE 记录不完整）")
            chunk = self.chunks[self.ci]
            take = min(count, len(chunk) - self.pos)
            out += chunk[self.pos:self.pos + take]
            self.pos += take
            count -= take
        return bytes(out)

    def skip(self, count: int) -> None:
        while count > 0:
            if self._at_end():
                raise ValueError("共享字符串表数据意外结束（CONTINUE 记录不完整）")
            chunk = self.chunks[self.ci]
            take = min(count, len(chunk) - self.pos)
            self.pos += take
            count -= take

    def u8(self) -> int:
        return self.read(1)[0]

    def u16(self) -> int:
        return struct.unpack("<H", self.read(2))[0]

    def u32(self) -> int:
        return struct.unpack("<I", self.read(4))[0]

    def chars(self, cch: int, high: bool) -> str:
        parts: list[str] = []
        remaining = cch
        while remaining > 0:
            if self.pos >= len(self.chunks[self.ci]):
                # 跨记录：CONTINUE 的首字节是新的 fHighByte 标志
                self.ci += 1
                self.pos = 0
                if self.ci >= len(self.chunks):
                    raise ValueError("共享字符串表数据意外结束（CONTINUE 记录不完整）")
                if not self.chunks[self.ci]:
                    continue
                high = bool(self.chunks[self.ci][0] & 0x01)
                self.pos = 1
            chunk = self.chunks[self.ci]
            unit = 2 if high else 1
            available = len(chunk) - self.pos
            take = min(remaining, available // unit)
            if take <= 0:
                if available:
                    raise ValueError("共享字符串跨记录边界落在半个字符上，无法解析")
                continue
            data = chunk[self.pos:self.pos + take * unit]
            self.pos += take * unit
            parts.append(data.decode("utf-16-le" if high else "latin-1", "ignore"))
            remaining -= take
        return "".join(parts)


def _read_unicode(body: bytes, off: int, lenbytes: int = 2) -> tuple[str, int]:
    """读取 BIFF8 XLUnicodeRichExtendedString（lenbytes=1 时为短串）。"""
    if lenbytes == 2:
        cch = _u16(body, off)
        off += 2
    else:
        if off >= len(body):
            raise ValueError("字符串长度字段越界")
        cch = body[off]
        off += 1
    if off >= len(body):
        raise ValueError("字符串标志字段越界")
    flags = body[off]
    off += 1
    high = flags & 0x01
    has_ext = flags & 0x04
    has_rich = flags & 0x08
    c_run = _u16(body, off) if has_rich else 0
    if has_rich:
        off += 2
    ext_len = _u32(body, off) if has_ext else 0
    if has_ext:
        off += 4
    unit = 2 if high else 1
    need = cch * unit
    if off + need > len(body):
        raise ValueError("字符串内容越界（记录被截断）")
    text = body[off:off + need].decode("utf-16-le" if high else "latin-1", "ignore")
    off += need + c_run * 4 + ext_len
    return text, off


def _rk_value(rk: int) -> float:
    """RK 压缩数值 → float（支持整数/浮点、×100 标志与 30 位有符号整数）。"""
    if rk & 0x02:
        value = rk >> 2
        if value & 0x20000000:  # 30 位补码的高位为符号位
            value -= 0x40000000
        number = float(value)
    else:
        number = struct.unpack("<d", struct.pack("<Q", (rk & 0xFFFFFFFC) << 32))[0]
    return number / 100.0 if rk & 0x01 else number


def _parse_sst(chunks: list[bytes], name: str) -> list[str]:
    if not chunks or len(chunks[0]) < 8:
        raise ValueError(f"{name} 的共享字符串表（SST）记录不完整")
    total, unique = struct.unpack_from("<II", chunks[0], 0)
    if unique > 5_000_000:
        raise ValueError(f"{name} 的共享字符串表数量异常：{unique}")
    reader = _ChunkReader([chunks[0][8:]] + chunks[1:])
    strings: list[str] = []
    try:
        for _ in range(unique):
            cch = reader.u16()
            flags = reader.u8()
            high = bool(flags & 0x01)
            has_ext = bool(flags & 0x04)
            has_rich = bool(flags & 0x08)
            c_run = reader.u16() if has_rich else 0
            ext_len = reader.u32() if has_ext else 0
            strings.append(reader.chars(cch, high))
            if c_run:
                reader.skip(c_run * 4)
            if ext_len:
                reader.skip(ext_len)
    except (ValueError, struct.error) as exc:
        raise ValueError(f"{name} 的共享字符串表（SST，声明 {unique} 条）解析失败：{exc}") from exc
    return strings


# ==========================================================================
# 单元格构造
# ==========================================================================
def _number_cell(value: float, ixfe: int,
                 style_dates: list[tuple[bool, bool] | None],
                 date1904: bool) -> Cell:
    if 0 <= ixfe < len(style_dates):
        style = style_dates[ixfe]
        if style is not None:
            text = format_serial_date(value, date1904, style[0], style[1])
            if text is not None:
                return Cell(raw=value, kind=KIND_DATE, text=text, num=value)
    raw = int(value) if value == int(value) and abs(value) < 1e16 else value
    return Cell(raw=raw, kind=KIND_NUMBER, text=_format_number(value), num=value)


def _formula_cell(result: bytes, ixfe: int,
                  style_dates: list[tuple[bool, bool] | None],
                  date1904: bool) -> tuple[Cell | None, bool]:
    """解析 FORMULA 的 8 字节缓存结果，返回 (cell, 是否等待后面的 STRING)。"""
    if len(result) < 8:
        raise ValueError("FORMULA 记录缓存值长度不足")
    if result[6] == 0xFF and result[7] == 0xFF:
        kind = result[0]
        if kind == 0:  # 字符串结果：紧随 STRING 记录
            return None, True
        if kind == 1:  # 布尔
            flag = result[2] != 0
            return Cell(raw=flag, kind=KIND_BOOL,
                        text="TRUE" if flag else "FALSE"), False
        if kind == 2:  # 错误
            text = _ERROR_TEXT.get(result[2], "#ERROR!")
            return Cell(raw=text, kind=KIND_TEXT, text=text), False
        return Cell(raw="", kind=KIND_FORMULA, text=""), False  # 空结果
    number = struct.unpack("<d", result[:8])[0]
    return _number_cell(number, ixfe, style_dates, date1904), False


# ==========================================================================
# 工作表解析
# ==========================================================================
def _parse_sheet(data: bytes, start: int, sheet_name: str,
                 sst: list[str],
                 style_dates: list[tuple[bool, bool] | None],
                 date1904: bool,
                 file_name: str) -> SheetModel:
    cells: dict[tuple[int, int], Cell] = {}
    merges: list[tuple[int, int, int, int]] = []
    max_r = max_c = -1
    depth = 0
    pending_string = False
    pending_pos: tuple[int, int] = (0, 0)

    def _touch(row: int, col: int) -> None:
        nonlocal max_r, max_c
        if row > max_r:
            max_r = row
        if col > max_c:
            max_c = col

    for rid, body in _iter_records(data, start):
        if rid == _RID_BOF:
            depth += 1
            if depth == 1 and len(body) >= 2:
                version = _u16(body, 0)
                if version != 0x0600:
                    raise ValueError(
                        f"{file_name} 的工作表 {sheet_name!r} 是 BIFF 0x{version:04X} 格式，"
                        "仅支持 BIFF8（Excel 97 及以上），请另存为 xlsx 后读取")
            continue
        if rid == _RID_EOF:
            depth -= 1
            if depth <= 0:
                break
            continue
        if pending_string and rid != _RID_STRING:
            pending_string = False
        try:
            if rid == _RID_LABELSST:
                row, col, ixfe, isst = struct.unpack_from("<HHHI", body, 0)
                _touch(row, col)
                if not (0 <= isst < len(sst)):
                    raise ValueError(
                        f"LABELSST 引用的共享字符串下标越界：{isst}（共 {len(sst)} 条）")
                text = sst[isst]
                cells[(row, col)] = Cell(raw=text, kind=KIND_TEXT, text=text)
            elif rid == _RID_LABEL:
                row, col, ixfe = struct.unpack_from("<HHH", body, 0)
                _touch(row, col)
                text, _ = _read_unicode(body, 6, 2)
                cells[(row, col)] = Cell(raw=text, kind=KIND_TEXT, text=text)
            elif rid == _RID_RSTRING:
                row, col, ixfe = struct.unpack_from("<HHH", body, 0)
                _touch(row, col)
                text, _ = _read_unicode(body, 6, 2)
                cells[(row, col)] = Cell(raw=text, kind=KIND_TEXT, text=text)
            elif rid == _RID_NUMBER:
                row, col, ixfe = struct.unpack_from("<HHH", body, 0)
                _touch(row, col)
                value = struct.unpack_from("<d", body, 6)[0]
                cells[(row, col)] = _number_cell(value, ixfe, style_dates, date1904)
            elif rid == _RID_RK:
                row, col, ixfe = struct.unpack_from("<HHH", body, 0)
                _touch(row, col)
                rk = _u32(body, 6)
                cells[(row, col)] = _number_cell(_rk_value(rk), ixfe,
                                                 style_dates, date1904)
            elif rid == _RID_MULRK:
                row, col_first = struct.unpack_from("<HH", body, 0)
                count = max(0, (len(body) - 6) // 6)
                for i in range(count):
                    offset = 4 + i * 6
                    ixfe = _u16(body, offset)
                    rk = _u32(body, offset + 2)
                    col = col_first + i
                    _touch(row, col)
                    cells[(row, col)] = _number_cell(_rk_value(rk), ixfe,
                                                     style_dates, date1904)
            elif rid == _RID_BOOLERR:
                row, col, ixfe = struct.unpack_from("<HHH", body, 0)
                _touch(row, col)
                value, is_error = body[6], body[7]
                if is_error:
                    text = _ERROR_TEXT.get(value, "#ERROR!")
                    cells[(row, col)] = Cell(raw=text, kind=KIND_TEXT, text=text)
                else:
                    cells[(row, col)] = Cell(raw=bool(value), kind=KIND_BOOL,
                                             text="TRUE" if value else "FALSE")
            elif rid == _RID_FORMULA:
                row, col, ixfe = struct.unpack_from("<HHH", body, 0)
                _touch(row, col)
                cell, waiting = _formula_cell(body[6:14], ixfe,
                                              style_dates, date1904)
                if waiting:
                    pending_string = True
                    pending_pos = (row, col)
                    cells[(row, col)] = Cell(raw=None, kind=KIND_FORMULA,
                                             text="=<公式>")
                elif cell is not None:
                    cells[(row, col)] = cell
            elif rid == _RID_STRING:
                if pending_string:
                    text, _ = _read_unicode(body, 0, 2)
                    row, col = pending_pos
                    cells[(row, col)] = Cell(raw=text, kind=KIND_TEXT, text=text)
                    pending_string = False
            elif rid == _RID_MERGEDCELLS:
                count = _u16(body, 0)
                if len(body) < 2 + count * 8:
                    raise ValueError("MERGEDCELLS 记录长度不足")
                for i in range(count):
                    rf, rl, cf, cl = struct.unpack_from("<HHHH", body, 2 + i * 8)
                    merges.append((rf, rl, cf, cl))
                    if rl > max_r:
                        max_r = rl
                    if cl > max_c:
                        max_c = cl
            elif rid == _RID_BLANK:
                # 仅表示"有格式的空单元格"：WPS 会把整列格式填充到 65536 行，
                # 不能计入使用范围，否则 n_rows 会被撑爆。
                pass
            elif rid == _RID_MULBLANK:
                pass
        except (struct.error, IndexError) as exc:
            raise ValueError(
                f"{file_name} 的工作表 {sheet_name!r} 中 BIFF 记录 "
                f"0x{rid:04X} 解析失败：{exc}") from exc

    if max_r < 0 and max_c < 0:
        # 完全是空表（可能只有格式）：按 1x1 空表返回，不采用可能是
        # 65536 的 DIMENSIONS 值，避免上层把空表当成超宽超高的表
        n_rows, n_cols = 1, 1
    else:
        n_rows, n_cols = max_r + 1, max_c + 1
    return SheetModel(name=sheet_name, n_rows=n_rows, n_cols=n_cols,
                      cells=cells, merges=merges)


# ==========================================================================
# 工作簿解析
# ==========================================================================
def _parse_workbook_stream(stream: bytes, wb: WorkbookModel, file_name: str) -> None:
    boundsheets: list[dict] = []
    sst: list[str] = []
    xf_formats: list[int] = []
    format_codes: dict[int, str] = {}
    date1904 = False
    version: int | None = None
    depth = 0

    pos = 0
    size = len(stream)
    while pos + 4 <= size:
        rid, length = struct.unpack_from("<HH", stream, pos)
        body = stream[pos + 4:pos + 4 + length]
        if rid == _RID_BOF:
            depth += 1
            if depth == 1:
                version = _u16(body, 0) if len(body) >= 2 else 0
                if version != 0x0600:
                    raise ValueError(
                        f"{file_name} 是 BIFF 0x{version:04X} 格式（Excel 95 及更早），"
                        "仅支持 BIFF8（Excel 97 及以上），请另存为 xlsx 后读取")
            pos += 4 + length
            continue
        if rid == _RID_EOF:
            depth -= 1
            if depth <= 0:
                break
            pos += 4 + length
            continue
        if rid == _RID_FILEPASS:
            raise ValueError(f"{file_name} 已加密，无法读取；请先在 Excel/WPS 中取消密码保护")
        if rid == _RID_BUNDLESHEET:
            offset = _u32(body, 0)
            hidden = bool(body[4] & 0x03)
            sheet_type = body[5]
            try:
                sheet_name, _ = _read_unicode(body, 6, 1)
            except ValueError:
                sheet_name = f"Sheet{len(boundsheets) + 1}"
            boundsheets.append({
                "offset": offset,
                "name": sheet_name or f"Sheet{len(boundsheets) + 1}",
                "hidden": hidden,
                "type": sheet_type,
            })
        elif rid == _RID_1904:
            date1904 = len(body) >= 2 and _u16(body, 0) == 1
        elif rid == _RID_XF:
            if len(body) >= 4:
                xf_formats.append(_u16(body, 2))
        elif rid == _RID_FORMAT:
            if len(body) >= 2:
                fmt_id = _u16(body, 0)
                try:
                    code, _ = _read_unicode(body, 2, 2)
                except ValueError:
                    code = ""
                format_codes[fmt_id] = code
        elif rid == _RID_SST:
            chunks = [body]
            next_pos = pos + 4 + length
            while next_pos + 4 <= size:
                next_rid, next_len = struct.unpack_from("<HH", stream, next_pos)
                if next_rid != _RID_CONTINUE:
                    break
                chunks.append(stream[next_pos + 4:next_pos + 4 + next_len])
                next_pos += 4 + next_len
            sst = _parse_sst(chunks, file_name)
            pos = next_pos
            continue
        pos += 4 + length

    if version is None:
        raise ValueError(f"{file_name} 的 Workbook 流没有 BOF 记录（文件损坏）")
    if not boundsheets:
        raise ValueError(f"{file_name} 的 Workbook 流里没有工作表（BOUNDSHEET 记录缺失）")

    style_dates = [classify_date_format(fmt, format_codes.get(fmt)) for fmt in xf_formats]
    for info in boundsheets:
        if info["type"] != 0:
            wb.notes.append(f"跳过非工作表对象：{info['name']}")
            continue
        sheet = _parse_sheet(stream, info["offset"], info["name"], sst,
                             style_dates, date1904, file_name)
        sheet.hidden = info["hidden"]
        wb.sheets.append(sheet)


# ==========================================================================
# 对外入口
# ==========================================================================
def read_biff(path: str) -> WorkbookModel:
    """读取 .xls / WPS .et（OLE 版）；失败抛 ValueError（中文原因）。"""
    path = os.fspath(path)
    file_name = os.path.basename(path)
    try:
        with open(path, "rb") as fp:
            data = fp.read()
    except OSError as exc:
        raise ValueError(f"无法读取文件：{path}（{exc.strerror or exc}）") from exc
    try:
        cfb = CFB(data, file_name)
    except ValueError:
        raise
    names = cfb.stream_names()
    stream_name = next((n for n in names if n.lower() == "workbook"), None)
    if stream_name is None:
        stream_name = next((n for n in names if n.lower() == "book"), None)
    if stream_name is None:
        raise ValueError(f"{file_name} 是 OLE2 复合文档，但其中没有 Workbook 流，"
                         "可能不是 Excel/WPS 表格文件（或为 Word 等文档）")
    try:
        stream = cfb.read_stream(stream_name)
    except (KeyError, ValueError) as exc:
        raise ValueError(f"{file_name} 的 Workbook 流读取失败：{exc}") from exc
    wb = WorkbookModel(path=path)
    _parse_workbook_stream(stream, wb, file_name)
    if not wb.sheets:
        raise ValueError(f"{file_name} 中没有可读取的工作表")
    return wb
