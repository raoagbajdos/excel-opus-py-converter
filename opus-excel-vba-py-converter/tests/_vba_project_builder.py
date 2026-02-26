"""
Build a valid vbaProject.bin (OLE2 / Compound Binary File) with embedded
VBA modules so that oletools / olevba can extract them.

Implements
  - MS-CFB   (Compound File Binary Format)  — the OLE2 container
  - MS-OVBA  (VBA File Format Specification) — dir stream, compression

Usage::

    from _vba_project_builder import build_vba_project

    modules = [
        ("Module1", 'Attribute VB_Name = "Module1"\\nSub Hi()\\nEnd Sub', False),
        ("clsPolicy", 'VERSION 1.0 CLASS\\n...', True),
    ]
    vba_bin = build_vba_project(modules)
    # Then inject vba_bin as  xl/vbaProject.bin  inside the .xlsm ZIP.
"""
from __future__ import annotations

import io
import struct


# ═══════════════════════════════════════════════════════════════════════════════
#  Constants
# ═══════════════════════════════════════════════════════════════════════════════

_SECTOR      = 512            # bytes per sector (CFB v3)
_ENDOFCHAIN  = 0xFFFFFFFE
_FREESECT    = 0xFFFFFFFF
_FATSECT     = 0xFFFFFFFD


# ═══════════════════════════════════════════════════════════════════════════════
#  1.  VBA source compression  (MS-OVBA §2.4.1 – literal-only variant)
# ═══════════════════════════════════════════════════════════════════════════════

def _compress(raw: bytes) -> bytes:
    """Compress *raw* using the MS-OVBA algorithm (literal tokens only).

    The output is valid for any compliant decompressor; it is simply not
    as small as it could be (no back-references / copy tokens).

    When literal-only compression would overflow the 12-bit chunk-size
    field (chunks of ~3641+ bytes), an *uncompressed* chunk is emitted
    instead (MS-OVBA §2.4.1.3.3).
    """
    out = bytearray(b"\x01")                    # container signature
    pos = 0
    while pos < len(raw):
        chunk = raw[pos : pos + 4096]
        pos += len(chunk)

        # — build token sequences (flag byte + up to 8 literal bytes) —
        tokens = bytearray()
        i = 0
        while i < len(chunk):
            n = min(8, len(chunk) - i)
            tokens.append(0x00)                  # all flag bits 0 → literal
            tokens.extend(chunk[i : i + n])
            i += n

        size_field = len(tokens) - 1

        if size_field > 0xFFF:
            # Literal-only compression inflated this chunk beyond the
            # 12-bit limit.  Emit an *uncompressed* chunk instead:
            #   - data is raw bytes padded to exactly 4096 with 0x00
            #   - CompressedChunkFlag = 0  (bit 15)
            #   - size field = 4096 + 2 - 3 = 4095 = 0x0FFF
            padded = chunk + b"\x00" * (4096 - len(chunk))
            hdr = 0x0FFF | (0b011 << 12)        # flag=0 (uncompressed)
            out.extend(struct.pack("<H", hdr))
            out.extend(padded)
        else:
            # — chunk header (2 bytes, little-endian) —
            # bits 11:0  = size_field
            # bits 14:12 = 0b011 (signature)
            # bit  15    = 1 (compressed)
            hdr = size_field | (0b011 << 12) | (1 << 15)
            out.extend(struct.pack("<H", hdr))
            out.extend(tokens)
    return bytes(out)


# ═══════════════════════════════════════════════════════════════════════════════
#  2.  dir stream builder  (MS-OVBA §2.3.4.2)
# ═══════════════════════════════════════════════════════════════════════════════

def _rec(buf: io.BytesIO, rid: int, data: bytes) -> None:
    """Write one TLV record: Id(2) + Size(4) + Data."""
    buf.write(struct.pack("<HI", rid, len(data)))
    buf.write(data)


def _build_dir_stream(modules: list[tuple[str, str, bool]]) -> bytes:
    """Return the **compressed** ``dir`` stream.

    Args:
        modules: ``[(name, vba_source, is_class_or_document), …]``
    """
    b = io.BytesIO()

    # ── Information records ──────────────────────────────────────────
    _rec(b, 0x0001, struct.pack("<I", 1))        # SysKind  Win32
    _rec(b, 0x0002, struct.pack("<I", 0x0409))   # LCID     en-US
    _rec(b, 0x0014, struct.pack("<I", 0x0409))   # LCIDInvoke
    _rec(b, 0x0003, struct.pack("<H", 1252))     # CodePage cp1252
    _rec(b, 0x0004, b"VBAProject")               # Name
    _rec(b, 0x0005, b"")                          # DocString
    _rec(b, 0x0040, b"")                          # DocString  (UTF-16)
    _rec(b, 0x0006, b"")                          # HelpFilePath1
    _rec(b, 0x003D, b"")                          # HelpFilePath2
    _rec(b, 0x0007, struct.pack("<I", 0))        # HelpContext
    _rec(b, 0x0008, struct.pack("<I", 0))        # LibFlags
    # PROJECTVERSION – spec §2.3.4.2.1.10: Reserved MUST be 0x00000004
    b.write(struct.pack("<HI", 0x0009, 0x0004))      # Id + Reserved
    b.write(struct.pack("<IH", 1467127604, 14))       # MajorVersion + MinorVersion
    _rec(b, 0x000C, b"")                          # Constants
    _rec(b, 0x003C, b"")                          # Constants  (UTF-16)

    # ── References (minimal – stdole only) ───────────────────────────
    _rec(b, 0x0016, b"stdole")                    # RefName
    libid = (
        b"*\\G{00020430-0000-0000-C000-000000000046}"
        b"#2.0#0#C:\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation"
    )
    _rec(b, 0x000D, struct.pack("<I", len(libid)) + libid + struct.pack("<IH", 0, 0))

    # ── Module section ───────────────────────────────────────────────
    _rec(b, 0x000F, struct.pack("<H", len(modules)))   # count
    _rec(b, 0x0013, struct.pack("<H", 0xFFFF))         # cookie

    for name, _src, is_class in modules:
        nb = name.encode("ascii")
        nu = name.encode("utf-16-le")
        _rec(b, 0x0019, nb)                   # ModuleName
        _rec(b, 0x0047, nu)                   # ModuleName  (UTF-16)
        _rec(b, 0x001A, nb)                   # StreamName
        _rec(b, 0x0032, nu)                   # StreamName  (UTF-16)
        _rec(b, 0x001C, b"")                  # DocString
        _rec(b, 0x0048, b"")                  # DocString  (UTF-16)
        _rec(b, 0x0031, struct.pack("<I", 0)) # Offset (source at byte 0)
        _rec(b, 0x001E, struct.pack("<I", 0)) # HelpContext
        _rec(b, 0x002C, struct.pack("<H", 0xFFFF))  # Cookie
        _rec(b, 0x0022 if is_class else 0x0021, b"")  # Type
        _rec(b, 0x002B, b"")                  # ModuleEnd

    return _compress(b.getvalue())


# ═══════════════════════════════════════════════════════════════════════════════
#  3.  Helper streams
# ═══════════════════════════════════════════════════════════════════════════════

def _build_project_stream(modules: list[tuple[str, str, bool]]) -> bytes:
    lines: list[str] = ['ID="{00000000-0000-0000-0000-000000000000}"']
    for name, _, is_cls in modules:
        if name == "ThisWorkbook":
            lines.append(f"Document={name}/&H00000000")
        elif is_cls:
            lines.append(f"Class={name}")
        else:
            lines.append(f"Module={name}")
    lines += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="0000"', 'DPB="0000"', 'GC="0000"', "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "", "[Workspace]",
    ]
    for name, _, _ in modules:
        lines.append(f"{name}=0, 0, 0, 0, C")
    return "\r\n".join(lines).encode("ascii")


def _build_projectwm_stream(modules: list[tuple[str, str, bool]]) -> bytes:
    buf = io.BytesIO()
    for name, _, _ in modules:
        buf.write(name.encode("ascii") + b"\x00")
        buf.write(name.encode("utf-16-le") + b"\x00\x00")
    buf.write(b"\x00")
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
#  4.  OLE2 / CFB container builder  (MS-CFB)
# ═══════════════════════════════════════════════════════════════════════════════

def _nsectors(size: int) -> int:
    """How many 512-byte sectors are needed for *size* bytes?"""
    return (size + _SECTOR - 1) // _SECTOR if size else 0


def _pad(data: bytes) -> bytes:
    """Pad *data* to a sector boundary."""
    r = len(data) % _SECTOR
    return data + b"\x00" * (_SECTOR - r) if r else data


def _balanced_tree(
    indices: list[int],
) -> tuple[int, dict[int, tuple[int, int]]]:
    """Return ``(root, {idx: (left, right)})`` for a balanced BST."""
    if not indices:
        return -1, {}
    mid = len(indices) // 2
    root = indices[mid]
    lr, lp = _balanced_tree(indices[:mid])
    rr, rp = _balanced_tree(indices[mid + 1 :])
    ptrs: dict[int, tuple[int, int]] = {root: (lr, rr)}
    ptrs.update(lp)
    ptrs.update(rp)
    return root, ptrs


class _Entry:
    """128-byte OLE2 directory entry."""

    __slots__ = ("name", "typ", "left", "right", "child", "start", "size")

    def __init__(self, name: str, typ: int, start: int = _ENDOFCHAIN, size: int = 0):
        self.name = name
        self.typ = typ          # 1=storage  2=stream  5=root
        self.left = -1
        self.right = -1
        self.child = -1
        self.start = start
        self.size = size

    def pack(self) -> bytes:
        buf = bytearray(128)
        nu = self.name.encode("utf-16-le")[:62]
        buf[: len(nu)] = nu
        struct.pack_into("<H", buf, 0x40, len(nu) + 2)     # name length
        buf[0x42] = self.typ
        buf[0x43] = 0x01                                     # colour = black
        struct.pack_into("<i", buf, 0x44, self.left)
        struct.pack_into("<i", buf, 0x48, self.right)
        struct.pack_into("<i", buf, 0x4C, self.child)
        struct.pack_into("<I", buf, 0x74, self.start & 0xFFFFFFFF)
        struct.pack_into("<I", buf, 0x78, self.size)
        return bytes(buf)


def _build_cfb(streams: list[tuple[str, bytes]]) -> bytes:
    """Build a minimal OLE2 compound-binary file.

    *streams* is an ordered list of ``("path", data)`` pairs.
    One level of storage nesting (``VBA/dir``) is supported.
    """
    # ── 1. Parse paths into Root / Storage grouping ──────────────────
    root_kids: list[tuple[str, bytes | None]] = []   # (name, data-or-None)
    vba_kids:  list[tuple[str, bytes]]         = []

    for path, data in streams:
        if path.startswith("VBA/"):
            vba_kids.append((path[4:], data))
        else:
            root_kids.append((path, data))

    has_vba = bool(vba_kids)

    # ── 2. Create directory entries ──────────────────────────────────
    entries: list[_Entry] = [_Entry("Root Entry", 5)]    # idx 0

    root_child_idx: list[int] = []
    vba_idx = -1

    if has_vba:
        vba_idx = len(entries)
        entries.append(_Entry("VBA", 1))
        root_child_idx.append(vba_idx)

    # Root-level streams
    root_stream_map: dict[int, bytes] = {}
    for name, data in root_kids:
        idx = len(entries)
        entries.append(_Entry(name, 2, size=len(data)))
        root_stream_map[idx] = data
        root_child_idx.append(idx)

    # VBA-level streams
    vba_child_idx: list[int] = []
    vba_stream_map: dict[int, bytes] = {}
    for name, data in vba_kids:
        idx = len(entries)
        entries.append(_Entry(name, 2, size=len(data)))
        vba_stream_map[idx] = data
        vba_child_idx.append(idx)

    all_stream_map = {**root_stream_map, **vba_stream_map}

    # ── 3. Wire up the balanced directory trees ──────────────────────
    root_child_idx.sort(key=lambda i: entries[i].name.upper())
    rt_root, rt_ptrs = _balanced_tree(root_child_idx)
    entries[0].child = rt_root
    for i, (l, r) in rt_ptrs.items():
        entries[i].left, entries[i].right = l, r

    if has_vba:
        vba_child_idx.sort(key=lambda i: entries[i].name.upper())
        vt_root, vt_ptrs = _balanced_tree(vba_child_idx)
        entries[vba_idx].child = vt_root
        for i, (l, r) in vt_ptrs.items():
            entries[i].left, entries[i].right = l, r

    # ── 4. Separate small / large streams ───────────────────────────
    MINI_CUTOFF = 0x1000  # 4096 – mandatory value in MS-CFB
    MINI_SECTOR = 64

    small: dict[int, bytes] = {}   # entry-idx → data  (mini stream)
    large: dict[int, bytes] = {}   # entry-idx → data  (regular sectors)

    for idx, blob in all_stream_map.items():
        if not blob:
            entries[idx].start = _ENDOFCHAIN
        elif len(blob) < MINI_CUTOFF:
            small[idx] = blob
        else:
            large[idx] = blob

    # ── 4a. Build mini-stream container and mini-FAT ────────────────
    mini_stream_data = bytearray()
    mini_fat: list[int] = []

    for idx in sorted(small):
        blob = small[idx]
        first_mini = len(mini_stream_data) // MINI_SECTOR
        entries[idx].start = first_mini

        # pad blob to mini-sector boundary
        padded_len = ((len(blob) + MINI_SECTOR - 1) // MINI_SECTOR) * MINI_SECTOR
        padded = blob + b"\x00" * (padded_len - len(blob))
        n_mini = padded_len // MINI_SECTOR

        for s in range(n_mini):
            mini_fat.append(first_mini + s + 1 if s < n_mini - 1 else _ENDOFCHAIN)
        mini_stream_data.extend(padded)

    # Pad mini-FAT to mini-FAT sector boundary (stored in regular sectors)
    has_mini = bool(mini_stream_data)

    # ── 5. Sector layout ────────────────────────────────────────────
    #   Sector 0         : FAT
    #   Sector 1..D      : Directory
    #   Sector D+1..M    : Mini-FAT sectors (if any)
    #   Sector M+1..R    : Mini-stream container (Root Entry data)
    #   Sector R+1..     : Large-stream data sectors
    dir_sectors = _nsectors(len(entries) * 128)

    mini_fat_bytes = b""
    if has_mini:
        # Mini-FAT: each entry = 4 bytes
        mf = bytearray(len(mini_fat) * 4)
        for i, v in enumerate(mini_fat):
            struct.pack_into("<I", mf, i * 4, v & 0xFFFFFFFF)
        # Pad to sector boundary
        mini_fat_bytes = _pad(bytes(mf))

    mini_fat_sectors = _nsectors(len(mini_fat_bytes)) if has_mini else 0
    mini_stream_sectors = _nsectors(len(mini_stream_data)) if has_mini else 0

    # Build FAT
    fat: list[int] = [_FATSECT]  # sector 0 = FAT

    # Directory chain: sectors 1 .. 1+dir_sectors-1
    for s in range(dir_sectors):
        fat.append(1 + s + 1 if s < dir_sectors - 1 else _ENDOFCHAIN)

    cur = 1 + dir_sectors  # next free sector

    # Mini-FAT chain
    first_mini_fat_sector = _ENDOFCHAIN
    if has_mini:
        first_mini_fat_sector = cur
        for s in range(mini_fat_sectors):
            fat.append(cur + s + 1 if s < mini_fat_sectors - 1 else _ENDOFCHAIN)
        cur += mini_fat_sectors

    # Mini-stream container (Root Entry data)
    if has_mini:
        entries[0].start = cur
        entries[0].size = len(mini_stream_data)
        for s in range(mini_stream_sectors):
            fat.append(cur + s + 1 if s < mini_stream_sectors - 1 else _ENDOFCHAIN)
        cur += mini_stream_sectors

    # Large streams
    data_blobs: list[bytes] = []
    for idx in sorted(large):
        blob = large[idx]
        entries[idx].start = cur
        ns = _nsectors(len(blob))
        for s in range(ns):
            fat.append(cur + s + 1 if s < ns - 1 else _ENDOFCHAIN)
        data_blobs.append(blob)
        cur += ns

    while len(fat) < _SECTOR // 4:
        fat.append(_FREESECT)

    # ── 6. Assemble binary ──────────────────────────────────────────
    out = bytearray()

    # — Header (512 bytes) —
    hdr = bytearray(_SECTOR)
    struct.pack_into("<8s", hdr, 0x00, b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
    struct.pack_into("<H",  hdr, 0x18, 0x003E)          # minor ver
    struct.pack_into("<H",  hdr, 0x1A, 0x0003)          # major ver 3
    struct.pack_into("<H",  hdr, 0x1C, 0xFFFE)          # byte order LE
    struct.pack_into("<H",  hdr, 0x1E, 9)               # sector pow 2^9
    struct.pack_into("<H",  hdr, 0x20, 6)               # mini-sector pow 2^6
    struct.pack_into("<I",  hdr, 0x2C, 1)               # FAT sectors = 1
    struct.pack_into("<I",  hdr, 0x30, 1)               # first dir sector
    struct.pack_into("<I",  hdr, 0x38, MINI_CUTOFF)     # mini cutoff = 4096
    struct.pack_into("<I",  hdr, 0x3C, first_mini_fat_sector & 0xFFFFFFFF)
    struct.pack_into("<I",  hdr, 0x40, mini_fat_sectors)
    struct.pack_into("<I",  hdr, 0x44, _ENDOFCHAIN)     # no DIFAT
    struct.pack_into("<I",  hdr, 0x48, 0)               # DIFAT count
    struct.pack_into("<I",  hdr, 0x4C, 0)               # DIFAT[0] → sector 0
    for d in range(1, 109):
        struct.pack_into("<I", hdr, 0x4C + d * 4, _FREESECT)
    out.extend(hdr)

    # — Sector 0: FAT —
    fat_sec = bytearray(_SECTOR)
    for fi, fv in enumerate(fat):
        struct.pack_into("<I", fat_sec, fi * 4, fv & 0xFFFFFFFF)
    out.extend(fat_sec)

    # — Directory sectors —
    dir_buf = bytearray()
    for e in entries:
        dir_buf.extend(e.pack())
    dir_buf += b"\x00" * (dir_sectors * _SECTOR - len(dir_buf))
    out.extend(dir_buf)

    # — Mini-FAT sectors —
    if has_mini:
        out.extend(mini_fat_bytes)

    # — Mini-stream container (Root Entry data) —
    if has_mini:
        out.extend(_pad(bytes(mini_stream_data)))

    # — Large-stream data sectors —
    for blob in data_blobs:
        out.extend(_pad(blob))

    return bytes(out)


# ═══════════════════════════════════════════════════════════════════════════════
#  5.  Public API
# ═══════════════════════════════════════════════════════════════════════════════

def build_vba_project(modules: list[tuple[str, str, bool]]) -> bytes:
    """Build a complete vbaProject.bin ready for an ``.xlsm`` archive.

    Args:
        modules: ``[(module_name, vba_source, is_class_or_document), …]``
                 Each *vba_source* should start with ``Attribute VB_Name``.

    Returns:
        The OLE2 binary (bytes) to write as ``xl/vbaProject.bin``.
    """
    streams: list[tuple[str, bytes]] = [
        ("VBA/dir",          _build_dir_stream(modules)),
        ("VBA/_VBA_PROJECT", b"\xCC\x61\xFF\xFF\x00\x00\x00"),
    ]
    for name, source, _ in modules:
        streams.append((f"VBA/{name}", _compress(source.encode("latin-1"))))
    streams.append(("PROJECT",   _build_project_stream(modules)))
    streams.append(("PROJECTwm", _build_projectwm_stream(modules)))

    return _build_cfb(streams)


# ═══════════════════════════════════════════════════════════════════════════════
#  6.  Quick self-test
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    _test_src = (
        'Attribute VB_Name = "Module1"\r\n'
        "Sub Hello()\r\n"
        '    MsgBox "Hello from VBA"\r\n'
        "End Sub\r\n"
    )
    data = build_vba_project([("Module1", _test_src, False)])
    assert data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", "bad OLE2 magic"
    print(f"vbaProject.bin: {len(data)} bytes – OLE2 magic OK")

    # Try opening with olefile
    import tempfile, olefile  # noqa: E401
    with tempfile.NamedTemporaryFile(suffix=".bin", delete=False) as f:
        f.write(data)
        tmp = f.name
    ole = olefile.OleFileIO(tmp)
    print("Streams:", ole.listdir())
    ole.close()
    import os; os.unlink(tmp)  # noqa: E702
    print("PASS ✓")
