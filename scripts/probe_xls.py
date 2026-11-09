# -*- coding: utf-8 -*-
"""Feasibility spike: pure-stdlib OLE2 + BIFF8 probe for legacy .xls files.

Goal: prove we can read .xls (the dominant format in the target folders)
without xlrd/LibreOffice, so the web app can stay dependency-free.
"""
import struct, sys, collections


class CFB:
    """Minimal OLE2 / Compound File Binary reader."""

    def __init__(self, data: bytes):
        self.d = data
        if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise ValueError("not an OLE2 file")
        self.ssz = 1 << struct.unpack_from("<H", data, 30)[0]
        self.mssz = 1 << struct.unpack_from("<H", data, 32)[0]
        (self.n_fat, self.first_dir, _, self.mini_cutoff,
         self.first_minifat, self.n_minifat, self.first_difat, self.n_difat) = struct.unpack_from(
            "<IIIIIIII", data, 44)
        self.fat = self._read_fat()
        self.dir = self._read_dir()
        self.minifat = self._read_minifat()

    def _sector(self, sid: int) -> bytes:
        off = (sid + 1) * self.ssz
        return self.d[off:off + self.ssz]

    def _read_fat(self):
        fat = []
        sid = self.first_difat
        fat_sids = []
        # header DIFAT has 109 entries
        for i in range(109):
            s = struct.unpack_from("<I", self.d, 76 + i * 4)[0]
            if s >= 0xFFFFFFFA:
                break
            fat_sids.append(s)
        while sid < 0xFFFFFFFA and len(fat_sids) < self.n_fat:
            sec = self._sector(sid)
            entries = struct.unpack("<%dI" % (self.ssz // 4 - 1), sec[:self.ssz - 4])
            fat_sids.extend(x for x in entries if x < 0xFFFFFFFA)
            sid = struct.unpack_from("<I", sec, self.ssz - 4)[0]
        for s in fat_sids:
            fat.extend(struct.unpack("<%dI" % (self.ssz // 4), self._sector(s)))
        return fat

    def _chain(self, start: int) -> list[int]:
        out, cur, seen = [], start, set()
        while cur < 0xFFFFFFFA and cur not in seen:
            seen.add(cur)
            out.append(cur)
            cur = self.fat[cur] if cur < len(self.fat) else 0xFFFFFFFE
        return out

    def _stream(self, start: int, size: int) -> bytes:
        buf = b"".join(self._sector(s) for s in self._chain(start))
        return buf[:size]

    def _read_minifat(self):
        if self.n_minifat == 0:
            return []
        raw = self._stream(self.first_minifat, self.n_minifat * self.ssz)
        return list(struct.unpack("<%dI" % (len(raw) // 4), raw))

    def _mini_chain(self, start: int) -> bytes:
        out, cur, seen = b"", start, set()
        while cur < 0xFFFFFFFA and cur not in seen:
            seen.add(cur)
            out += self.ministream[cur * self.mssz:(cur + 1) * self.mssz]
            cur = self.minifat[cur] if cur < len(self.minifat) else 0xFFFFFFFE
        return out

    def _read_dir(self):
        entries = []
        raw = self._stream(self.first_dir, self.ssz)
        # directory may span sectors
        n = self.ssz // 128
        for k in range(n):
            e = raw[k * 128:(k + 1) * 128]
            if len(e) < 128:
                break
            name_len = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(0, name_len - 2)].decode("utf-16-le", "ignore")
            etype = e[66]
            start = struct.unpack_from("<I", e, 116)[0]
            size = struct.unpack_from("<Q", e, 120)[0]
            entries.append({"name": name, "type": etype, "start": start, "size": size})
        # continue over remaining directory sectors
        cur = self.fat[self.first_dir] if self.first_dir < len(self.fat) else 0xFFFFFFFE
        while cur < 0xFFFFFFFA:
            raw = self._sector(cur)
            for k in range(n):
                e = raw[k * 128:(k + 1) * 128]
                name_len = struct.unpack_from("<H", e, 64)[0]
                name = e[:max(0, name_len - 2)].decode("utf-16-le", "ignore")
                etype = e[66]
                start = struct.unpack_from("<I", e, 116)[0]
                size = struct.unpack_from("<Q", e, 120)[0]
                entries.append({"name": name, "type": etype, "start": start, "size": size})
            cur = self.fat[cur] if cur < len(self.fat) else 0xFFFFFFFE
        return entries

    def read_stream(self, name: str) -> bytes:
        root = next(e for e in self.dir if e["type"] == 5)
        ministream = self._stream(root["start"], root["size"]) if root["start"] < 0xFFFFFFFA else b""
        self.ministream = ministream
        for e in self.dir:
            if e["type"] == 2 and e["name"] == name:
                if e["size"] < self.mini_cutoff:
                    return self._mini_chain(e["start"])[:e["size"]]
                return self._stream(e["start"], e["size"])
        raise KeyError(name)


# ---- BIFF8 record parsing -------------------------------------------------
def records(data: bytes):
    pos = 0
    while pos + 4 <= len(data):
        rid, ln = struct.unpack_from("<HH", data, pos)
        body = data[pos + 4:pos + 4 + ln]
        yield rid, body
        pos += 4 + ln


def read_unicode(body: bytes, off: int, lenbytes: int = 2):
    """Read a BIFF8 unicode string starting at off; return (text, new_off)."""
    if lenbytes == 2:
        cch = struct.unpack_from("<H", body, off)[0]
        off += 2
    else:
        cch = body[off]
        off += 1
    flags = body[off]
    off += 1
    high = flags & 0x01
    ext = flags & 0x04
    rich = flags & 0x08
    c_run = struct.unpack_from("<H", body, off)[0] if rich else 0
    off += 2 if rich else 0
    ext_len = struct.unpack_from("<I", body, off)[0] if ext else 0
    off += 4 if ext else 0
    if high:
        text = body[off:off + cch * 2].decode("utf-16-le", "ignore")
        off += cch * 2
    else:
        text = body[off:off + cch].decode("cp1252", "ignore")
        off += cch
    off += c_run * 4 + ext_len
    return text, off


def rk_value(rk: int):
    if rk & 0x02:
        val = float(rk >> 2)
    else:
        val = struct.unpack("<d", struct.pack("<Q", (rk & 0xFFFFFFFC) << 32))[0]
    return val / 100.0 if rk & 0x01 else val


def probe(path: str):
    data = open(path, "rb").read()
    cfb = CFB(data)
    stream_name = "Workbook" if any(e["name"] == "Workbook" for e in cfb.dir) else "Book"
    wb = cfb.read_stream(stream_name)
    sst, sheets, sheet_offsets = [], [], []
    for rid, body in records(wb):
        if rid == 0x0085:  # BOUNDSHEET
            off = struct.unpack_from("<I", body, 0)[0]
            hidden = body[4]
            typ = body[5]
            name, _ = read_unicode(body, 6, lenbytes=1)
            sheets.append((name, off, typ, hidden))
        elif rid == 0x00FC:  # SST
            total = struct.unpack_from("<I", body, 0)[0]
            uniq = struct.unpack_from("<I", body, 4)[0]
            off = 8
            count = 0
            while count < uniq and off < len(body):
                try:
                    text, off = read_unicode(body, off, lenbytes=2)
                except Exception:
                    break
                sst.append(text)
                count += 1
        elif rid == 0x000A:  # EOF of globals
            break

    results = []
    for (name, off, typ, hidden) in sheets:
        sub = wb[off:]
        cells = collections.OrderedDict()
        merged = []
        maxr = maxc = 0
        for rid, body in records(sub):
            try:
                if rid == 0x000A:
                    break
                if rid == 0x00FD:  # LABELSST
                    r, c, _x, idx = struct.unpack_from("<HHHI", body, 0)
                    cells[(r, c)] = sst[idx] if idx < len(sst) else ""
                elif rid == 0x0203:  # NUMBER
                    r, c, _x = struct.unpack_from("<HHH", body, 0)
                    v = struct.unpack_from("<d", body, 6)[0]
                    cells[(r, c)] = v
                elif rid == 0x027E:  # RK
                    r, c, _x = struct.unpack_from("<HHH", body, 0)
                    cells[(r, c)] = rk_value(struct.unpack_from("<I", body, 6)[0])
                elif rid == 0x00BD:  # MULRK
                    r, c0 = struct.unpack_from("<HH", body, 0)
                    n = (len(body) - 6) // 6
                    for i in range(n):
                        rk = struct.unpack_from("<I", body, 4 + i * 6 + 2)[0]
                        cells[(r, c0 + i)] = rk_value(rk)
                elif rid == 0x0204:  # LABEL (inline string)
                    r, c, _x = struct.unpack_from("<HHH", body, 0)
                    text, _ = read_unicode(body, 6, lenbytes=2)
                    cells[(r, c)] = text
                elif rid == 0x0205:  # BOOLERR
                    r, c, _x = struct.unpack_from("<HHH", body, 0)
                    cells[(r, c)] = bool(body[6])
                elif rid == 0x0006:  # FORMULA
                    r, c, _x = struct.unpack_from("<HHH", body, 0)
                    cells[(r, c)] = "<formula>"
                elif rid == 0x0207:  # STRING (result of formula)
                    text, _ = read_unicode(body, 0, lenbytes=2)
                    for (rr, cc) in reversed(list(cells.keys())):
                        if cells[(rr, cc)] == "<formula>":
                            cells[(rr, cc)] = text
                            break
                elif rid == 0x00E5:  # MERGEDCELLS
                    n = struct.unpack_from("<H", body, 0)[0]
                    for i in range(n):
                        rf, rl, cf, cl = struct.unpack_from("<HHHH", body, 2 + i * 8)
                        merged.append((rf, rl, cf, cl))
                for (r, c) in cells:
                    maxr = max(maxr, r)
                    maxc = max(maxc, c)
            except Exception:
                continue
        results.append((name, typ, merged, cells, maxr, maxc))
    return sst, results


if __name__ == "__main__":
    path = sys.argv[1]
    sst, sheets = probe(path)
    print(f"FILE {path}")
    print(f"SST strings: {len(sst)}")
    for name, typ, merged, cells, maxr, maxc in sheets:
        print("=" * 80)
        print(f"SHEET {name!r} type={typ} cells={len(cells)} merged={len(merged)} max=({maxr+1},{maxc+1})")
        print(" MERGED(first10):", merged[:10])
        for r in range(min(maxr + 1, 6)):
            row = [str(cells.get((r, c), ""))[:28] for c in range(min(maxc + 1, 12))]
            print(f" R{r+1}: " + " | ".join(row))
